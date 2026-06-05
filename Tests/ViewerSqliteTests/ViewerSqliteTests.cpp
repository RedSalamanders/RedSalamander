#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cmath>
#include <cstdlib>
#include <cstring>
#include <filesystem>
#include <format>
#include <iostream>
#include <memory>
#include <span>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>
#include <vector>

#include <UIAutomation.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#pragma warning(push)
#pragma warning(disable : 6297 28182)
#include <sqlite3.h>
#pragma warning(pop)

#include "PlugInterfaces/Factory.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/Viewer.h"
#include "ViewerSqlite.Engine.h"
#include "WindowMessages.h"

namespace
{
using unique_sqlite3 = std::unique_ptr<sqlite3, decltype(&sqlite3_close_v2)>;
using namespace std::chrono_literals;

constexpr wchar_t kViewerSqliteWindowClassName[] = L"RedSalamander.ViewerSqlite";

using RedSalamanderCreateFn = HRESULT(__stdcall*)(REFIID riid, const FactoryOptions* factoryOptions, IHost* host, const wchar_t* pluginId, void** result);
constexpr wchar_t kViewerSqlitePluginId[] = L"builtin/viewer-sqlite";

struct TempDatabase final
{
    std::filesystem::path path;

    TempDatabase()                                   = default;
    TempDatabase(const TempDatabase&)                = delete;
    TempDatabase& operator=(const TempDatabase&)     = delete;
    TempDatabase(TempDatabase&&) noexcept            = default;
    TempDatabase& operator=(TempDatabase&&) noexcept = default;

    ~TempDatabase() noexcept
    {
        if (path.empty())
        {
            return;
        }

        std::error_code ec;
        std::filesystem::remove(path, ec);
    }
};

struct ViewerClosedCounter final : IViewerCallback
{
    std::atomic<unsigned int> closedCount = 0u;

    ViewerClosedCounter()                                      = default;
    ViewerClosedCounter(const ViewerClosedCounter&)            = delete;
    ViewerClosedCounter(ViewerClosedCounter&&)                 = delete;
    ViewerClosedCounter& operator=(const ViewerClosedCounter&) = delete;
    ViewerClosedCounter& operator=(ViewerClosedCounter&&)      = delete;
    ~ViewerClosedCounter()                                     = default;

    HRESULT STDMETHODCALLTYPE ViewerClosed(void* cookie) noexcept override
    {
        if (cookie != this)
        {
            return E_INVALIDARG;
        }

        closedCount.fetch_add(1u, std::memory_order_relaxed);
        return S_OK;
    }
};

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text)
{
    if (text.empty())
    {
        return {};
    }

    const int needed = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (needed <= 0)
    {
        return {};
    }

    std::wstring out(static_cast<size_t>(needed), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), out.data(), needed);
    if (written != needed)
    {
        return {};
    }

    return out;
}

[[nodiscard]] bool Exec(sqlite3* db, const char* sql, std::wstring& errorText)
{
    char* err    = nullptr;
    const int rc = sqlite3_exec(db, sql, nullptr, nullptr, &err);
    if (rc == SQLITE_OK)
    {
        return true;
    }

    const std::string_view message = (err != nullptr) ? std::string_view(err) : std::string_view(sqlite3_errstr(rc));
    errorText                      = Utf16FromUtf8(message);
    sqlite3_free(err);
    return false;
}

void PumpPendingMessages() noexcept
{
    MSG msg{};
    while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) != 0)
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
}

[[nodiscard]] constexpr uint32_t Argb(uint8_t r, uint8_t g, uint8_t b) noexcept
{
    return 0xFF000000u | (static_cast<uint32_t>(r) << 16u) | (static_cast<uint32_t>(g) << 8u) | static_cast<uint32_t>(b);
}

#ifdef _DEBUG
[[nodiscard]] ViewerTheme MakeViewerTheme(BOOL darkMode,
                                          BOOL highContrast,
                                          BOOL rainbowMode,
                                          uint32_t backgroundArgb,
                                          uint32_t textArgb,
                                          uint32_t selectionBackgroundArgb,
                                          uint32_t selectionTextArgb,
                                          uint32_t accentArgb) noexcept
{
    ViewerTheme theme{};
    theme.version                    = 2u;
    theme.dpi                        = USER_DEFAULT_SCREEN_DPI;
    theme.backgroundArgb             = backgroundArgb;
    theme.textArgb                   = textArgb;
    theme.selectionBackgroundArgb    = selectionBackgroundArgb;
    theme.selectionTextArgb          = selectionTextArgb;
    theme.accentArgb                 = accentArgb;
    theme.alertErrorBackgroundArgb   = Argb(0x8B, 0x00, 0x00);
    theme.alertErrorTextArgb         = Argb(0xFF, 0xFF, 0xFF);
    theme.alertWarningBackgroundArgb = Argb(0xFF, 0xE0, 0x8A);
    theme.alertWarningTextArgb       = Argb(0x20, 0x20, 0x20);
    theme.alertInfoBackgroundArgb    = Argb(0xD8, 0xEC, 0xFF);
    theme.alertInfoTextArgb          = Argb(0x10, 0x10, 0x10);
    theme.darkMode                   = darkMode;
    theme.highContrast               = highContrast;
    theme.rainbowMode                = rainbowMode;
    theme.darkBase                   = darkMode;
    return theme;
}
#endif

[[nodiscard]] std::vector<HWND> CollectVisibleWindowsByClass(std::wstring_view className) noexcept
{
    std::vector<HWND> windows;
    struct EnumArgs
    {
        std::wstring_view className;
        std::vector<HWND>* windows = nullptr;
    } args{className, &windows};

    static_cast<void>(EnumWindows(
        [](HWND hwnd, LPARAM lParam) noexcept -> BOOL
    {
        auto* args = reinterpret_cast<EnumArgs*>(lParam);
        if (! args || IsWindowVisible(hwnd) == FALSE)
        {
            return TRUE;
        }

        wchar_t classBuffer[128] = {};
        const int classLength    = GetClassNameW(hwnd, classBuffer, static_cast<int>(std::size(classBuffer)));
        if (classLength <= 0)
        {
            return TRUE;
        }

        const std::wstring_view currentClass(classBuffer, static_cast<size_t>(classLength));
        if (currentClass.starts_with(args->className))
        {
            args->windows->push_back(hwnd);
        }
        return TRUE;
    },
        reinterpret_cast<LPARAM>(&args)));
    return windows;
}

[[nodiscard]] HWND FindNewVisibleWindowByClass(std::wstring_view className, std::span<const HWND> existingWindows) noexcept
{
    const std::vector<HWND> currentWindows = CollectVisibleWindowsByClass(className);
    const auto isExisting = [&](HWND hwnd) noexcept { return std::find(existingWindows.begin(), existingWindows.end(), hwnd) != existingWindows.end(); };

    const auto it = std::find_if(currentWindows.begin(), currentWindows.end(), [&](HWND hwnd) noexcept { return ! isExisting(hwnd); });
    return (it != currentWindows.end()) ? *it : nullptr;
}

template <typename Predicate> [[nodiscard]] bool PumpUntil(Predicate&& predicate, std::chrono::milliseconds timeout) noexcept
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (predicate())
        {
            return true;
        }
        std::this_thread::sleep_for(20ms);
    }

    PumpPendingMessages();
    return predicate();
}

[[nodiscard]] size_t CountVisibleChildWindows(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return 0u;
    }

    struct VisibleChildCounter
    {
        size_t count = 0u;
    } counter{};

    static_cast<void>(EnumChildWindows(hwnd,
                                       [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& counterRef = *reinterpret_cast<VisibleChildCounter*>(lParam);
        if (IsWindowVisible(child) != FALSE)
        {
            ++counterRef.count;
        }
        return TRUE;
    },
                                       reinterpret_cast<LPARAM>(&counter)));

    return counter.count;
}

#ifdef _DEBUG
[[nodiscard]] bool TryGetViewerSqliteDebugSnapshot(HWND hwnd, WndMsg::ViewerSqliteDebugSnapshot& snapshot) noexcept
{
    snapshot = {};
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    auto remoteSnapshot = std::make_unique<WndMsg::ViewerSqliteDebugSnapshot>();
    if (! remoteSnapshot)
    {
        return false;
    }

    if (SendMessageW(hwnd, WndMsg::kViewerSqliteDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(remoteSnapshot.get())) == FALSE)
    {
        return false;
    }

    snapshot = *remoteSnapshot;
    return true;
}

[[nodiscard]] bool DebugScrollViewerSqliteGridByWheelDetents(HWND hwnd, int detents) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    return SendMessageW(hwnd, WndMsg::kViewerSqliteDebugScrollGridByWheelDetents, static_cast<WPARAM>(detents), 0) != FALSE;
}

[[nodiscard]] bool DebugSelectViewerSqliteGridRow(HWND hwnd, size_t rowIndex) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    return SendMessageW(hwnd, WndMsg::kViewerSqliteDebugSelectGridRow, static_cast<WPARAM>(rowIndex), 0) != FALSE;
}

[[nodiscard]] bool DebugInvokeViewerSqlitePageCommand(HWND hwnd, WndMsg::ViewerSqliteDebugPageCommand command) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    return SendMessageW(hwnd, WndMsg::kViewerSqliteDebugInvokePageCommand, static_cast<WPARAM>(command), 0) != FALSE;
}

[[nodiscard]] bool DebugCycleViewerSqliteSortColumn(HWND hwnd, size_t columnIndex) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    return SendMessageW(hwnd, WndMsg::kViewerSqliteDebugCycleSortColumn, static_cast<WPARAM>(columnIndex), 0) != FALSE;
}

void SendViewerSqliteTab(HWND hwnd, bool reverse) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return;
    }

    if (reverse)
    {
        static_cast<void>(SendMessageW(hwnd, WM_KEYDOWN, VK_SHIFT, 0));
        static_cast<void>(SendMessageW(hwnd, WM_KEYDOWN, VK_TAB, 0));
        static_cast<void>(SendMessageW(hwnd, WM_KEYUP, VK_SHIFT, 0));
        return;
    }

    static_cast<void>(SendMessageW(hwnd, WM_KEYDOWN, VK_TAB, 0));
}

template <typename Predicate>
[[nodiscard]] bool WaitForViewerSnapshot(HWND hwnd,
                                         Predicate&& predicate,
                                         std::chrono::milliseconds timeout,
                                         WndMsg::ViewerSqliteDebugSnapshot* outSnapshot = nullptr) noexcept
{
    WndMsg::ViewerSqliteDebugSnapshot snapshot{};
    WndMsg::ViewerSqliteDebugSnapshot lastSnapshot{};
    bool sawSnapshot = false;
    const bool ready = PumpUntil(
        [&]() noexcept
    {
        if (! TryGetViewerSqliteDebugSnapshot(hwnd, snapshot))
        {
            return false;
        }

        lastSnapshot = snapshot;
        sawSnapshot  = true;
        if (! predicate(snapshot))
        {
            return false;
        }

        if (outSnapshot)
        {
            *outSnapshot = snapshot;
        }
        return true;
    },
        timeout);

    if (! ready && outSnapshot && sawSnapshot)
    {
        *outSnapshot = lastSnapshot;
    }

    return ready;
}
#endif

#if defined(_DEBUG) && ! defined(__SANITIZE_ADDRESS__)
struct UiaSelectionState
{
    size_t selectedCount              = 0u;
    CONTROLTYPEID selectedControlType = 0;
    std::wstring selectedName;
    bool selectedHasSelectionItemPattern = false;
    bool selectedVisible                 = false;
};
#endif

#ifdef _DEBUG
struct UiaViewerSubtreeStats
{
    size_t visibleElementCount   = 0u;
    size_t comboBoxControlCount  = 0u;
    size_t editControlCount      = 0u;
    size_t buttonControlCount    = 0u;
    size_t dataGridControlCount  = 0u;
    size_t dataItemControlCount  = 0u;
    size_t valuePatternCount     = 0u;
    size_t selectionPatternCount = 0u;
};

[[nodiscard]] bool TryCollectVisibleUiaViewerSubtreeStats(HWND hwnd, UiaViewerSubtreeStats& stats) noexcept
{
    stats = {};
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(coinitHr) && coinitHr != RPC_E_CHANGED_MODE)
    {
        return false;
    }

    const bool shouldUninitialize = SUCCEEDED(coinitHr);
    const auto coUninitialize     = wil::scope_exit([shouldUninitialize]() noexcept
    {
        if (shouldUninitialize)
        {
            CoUninitialize();
        }
    });

    wil::com_ptr<IUIAutomation> automation;
    HRESULT hr = CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(automation.addressof()));
    if (FAILED(hr) || ! automation)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationElement> root;
    hr = automation->ElementFromHandle(hwnd, root.addressof());
    if (FAILED(hr) || ! root)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationCondition> trueCondition;
    hr = automation->CreateTrueCondition(trueCondition.addressof());
    if (FAILED(hr) || ! trueCondition)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationElementArray> elements;
    hr = root->FindAll(TreeScope_Subtree, trueCondition.get(), elements.addressof());
    if (FAILED(hr) || ! elements)
    {
        return false;
    }

    int length = 0;
    hr         = elements->get_Length(&length);
    if (FAILED(hr) || length <= 0)
    {
        return false;
    }

    for (int index = 0; index < length; ++index)
    {
        wil::com_ptr<IUIAutomationElement> element;
        if (FAILED(elements->GetElement(index, element.addressof())) || ! element)
        {
            continue;
        }

        BOOL offscreen = TRUE;
        if (FAILED(element->get_CurrentIsOffscreen(&offscreen)) || offscreen != FALSE)
        {
            continue;
        }

        ++stats.visibleElementCount;

        CONTROLTYPEID controlType = 0;
        if (SUCCEEDED(element->get_CurrentControlType(&controlType)))
        {
            switch (controlType)
            {
                case UIA_ComboBoxControlTypeId: ++stats.comboBoxControlCount; break;
                case UIA_EditControlTypeId: ++stats.editControlCount; break;
                case UIA_ButtonControlTypeId: ++stats.buttonControlCount; break;
                case UIA_DataGridControlTypeId: ++stats.dataGridControlCount; break;
                case UIA_DataItemControlTypeId: ++stats.dataItemControlCount; break;
                default: break;
            }
        }

        wil::com_ptr<IUIAutomationValuePattern> valuePattern;
        if (SUCCEEDED(element->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), valuePattern.put_void())) && valuePattern)
        {
            ++stats.valuePatternCount;
        }

        wil::com_ptr<IUIAutomationSelectionPattern> selectionPattern;
        if (SUCCEEDED(element->GetCurrentPatternAs(UIA_SelectionPatternId, __uuidof(IUIAutomationSelectionPattern), selectionPattern.put_void())) &&
            selectionPattern)
        {
            ++stats.selectionPatternCount;
        }
    }

    return true;
}
#endif

#if defined(_DEBUG) && ! defined(__SANITIZE_ADDRESS__)
[[nodiscard]] bool TryCollectVisibleUiaViewerGridSelectionState(HWND hwnd, UiaSelectionState& state) noexcept
{
    state = {};
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(coinitHr) && coinitHr != RPC_E_CHANGED_MODE)
    {
        return false;
    }

    const bool shouldUninitialize = SUCCEEDED(coinitHr);
    const auto coUninitialize     = wil::scope_exit([shouldUninitialize]() noexcept
    {
        if (shouldUninitialize)
        {
            CoUninitialize();
        }
    });

    wil::com_ptr<IUIAutomation> automation;
    HRESULT hr = CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(automation.addressof()));
    if (FAILED(hr) || ! automation)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationElement> root;
    hr = automation->ElementFromHandle(hwnd, root.addressof());
    if (FAILED(hr) || ! root)
    {
        return false;
    }

    VARIANT controlTypeVariant{};
    V_VT(&controlTypeVariant) = VT_I4;
    V_I4(&controlTypeVariant) = UIA_DataGridControlTypeId;

    wil::com_ptr<IUIAutomationCondition> gridCondition;
    hr = automation->CreatePropertyCondition(UIA_ControlTypePropertyId, controlTypeVariant, gridCondition.addressof());
    if (FAILED(hr) || ! gridCondition)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationElement> grid;
    hr = root->FindFirst(TreeScope_Subtree, gridCondition.get(), grid.addressof());
    if (FAILED(hr) || ! grid)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationSelectionPattern> selectionPattern;
    hr = grid->GetCurrentPatternAs(UIA_SelectionPatternId, __uuidof(IUIAutomationSelectionPattern), selectionPattern.put_void());
    if (FAILED(hr) || ! selectionPattern)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationElementArray> selected;
    hr = selectionPattern->GetCurrentSelection(selected.addressof());
    if (FAILED(hr) || ! selected)
    {
        return false;
    }

    int selectedLength = 0;
    hr                 = selected->get_Length(&selectedLength);
    if (FAILED(hr) || selectedLength < 0)
    {
        return false;
    }

    state.selectedCount = static_cast<size_t>(selectedLength);
    if (selectedLength <= 0)
    {
        return true;
    }

    wil::com_ptr<IUIAutomationElement> selectedElement;
    hr = selected->GetElement(0, selectedElement.addressof());
    if (FAILED(hr) || ! selectedElement)
    {
        return false;
    }

    static_cast<void>(selectedElement->get_CurrentControlType(&state.selectedControlType));

    wil::unique_bstr name;
    if (SUCCEEDED(selectedElement->get_CurrentName(&name)))
    {
        state.selectedName.assign(name.get() ? name.get() : L"");
    }

    wil::com_ptr<IUIAutomationSelectionItemPattern> selectionItemPattern;
    if (SUCCEEDED(
            selectedElement->GetCurrentPatternAs(UIA_SelectionItemPatternId, __uuidof(IUIAutomationSelectionItemPattern), selectionItemPattern.put_void())) &&
        selectionItemPattern)
    {
        state.selectedHasSelectionItemPattern = true;
    }

    BOOL offscreen = TRUE;
    if (SUCCEEDED(selectedElement->get_CurrentIsOffscreen(&offscreen)))
    {
        state.selectedVisible = (offscreen == FALSE);
    }

    return true;
}
#endif

class BuiltinFileSystemStub final : public IFileSystem, public IInformations
{
public:
    BuiltinFileSystemStub()                                        = default;
    BuiltinFileSystemStub(const BuiltinFileSystemStub&)            = delete;
    BuiltinFileSystemStub(BuiltinFileSystemStub&&)                 = delete;
    BuiltinFileSystemStub& operator=(const BuiltinFileSystemStub&) = delete;
    BuiltinFileSystemStub& operator=(BuiltinFileSystemStub&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }

        *ppvObject = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
        {
            *ppvObject = static_cast<IFileSystem*>(this);
        }
        else if (riid == __uuidof(IInformations))
        {
            *ppvObject = static_cast<IInformations*>(this);
        }
        else
        {
            return E_NOINTERFACE;
        }

        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return static_cast<ULONG>(_refCount.fetch_add(1u, std::memory_order_relaxed) + 1u);
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        ULONG current = _refCount.load(std::memory_order_acquire);
        while (current != 0u)
        {
            if (_refCount.compare_exchange_weak(current, current - 1u, std::memory_order_acq_rel, std::memory_order_acquire))
            {
                return current - 1u;
            }
        }

        return 0u;
    }

    HRESULT STDMETHODCALLTYPE GetMetaData(const PluginMetaData** metaData) noexcept override
    {
        if (! metaData)
        {
            return E_POINTER;
        }

        *metaData = &kMetaData;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetConfigurationSchema(const char** schemaJsonUtf8) noexcept override
    {
        if (! schemaJsonUtf8)
        {
            return E_POINTER;
        }

        *schemaJsonUtf8 = nullptr;
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE SetConfiguration(const char* /*configurationJsonUtf8*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE GetConfiguration(const char** configurationJsonUtf8) noexcept override
    {
        if (! configurationJsonUtf8)
        {
            return E_POINTER;
        }

        *configurationJsonUtf8 = nullptr;
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE GetTransferHints([[maybe_unused]] const wchar_t* path,
                                               [[maybe_unused]] FileSystemOperation operationType,
                                               [[maybe_unused]] FileSystemTransferEndpoint endpoint,
                                               FileSystemTransferHints* hints) noexcept override
    {
        if (! path || path[0] == L'\0' || ! hints)
        {
            return E_INVALIDARG;
        }
        if (hints->sizeBytes < sizeof(FileSystemTransferHints))
        {
            return E_INVALIDARG;
        }
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE GetStorageCharacteristics([[maybe_unused]] const wchar_t* path,
                                                        FileSystemStorageCharacteristics* characteristics) noexcept override
    {
        if (! path || path[0] == L'\0' || ! characteristics)
        {
            return E_INVALIDARG;
        }
        if (characteristics->sizeBytes < sizeof(FileSystemStorageCharacteristics))
        {
            return E_INVALIDARG;
        }
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE SomethingToSave(BOOL* pSomethingToSave) noexcept override
    {
        if (! pSomethingToSave)
        {
            return E_POINTER;
        }

        *pSomethingToSave = FALSE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* /*path*/, IFilesInformation** /*ppFilesInformation*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* /*sourcePath*/,
                                       const wchar_t* /*destinationPath*/,
                                       FileSystemFlags /*flags*/,
                                       const FileSystemOptions* /*options*/,
                                       IFileSystemCallback* /*callback*/,
                                       void* /*cookie*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* /*sourcePath*/,
                                       const wchar_t* /*destinationPath*/,
                                       FileSystemFlags /*flags*/,
                                       const FileSystemOptions* /*options*/,
                                       IFileSystemCallback* /*callback*/,
                                       void* /*cookie*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE DeleteItem(const wchar_t* /*path*/,
                                         FileSystemFlags /*flags*/,
                                         const FileSystemOptions* /*options*/,
                                         IFileSystemCallback* /*callback*/,
                                         void* /*cookie*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* /*sourcePath*/,
                                         const wchar_t* /*destinationPath*/,
                                         FileSystemFlags /*flags*/,
                                         const FileSystemOptions* /*options*/,
                                         IFileSystemCallback* /*callback*/,
                                         void* /*cookie*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* /*sourcePaths*/,
                                        unsigned long /*count*/,
                                        const wchar_t* /*destinationFolder*/,
                                        FileSystemFlags /*flags*/,
                                        const FileSystemOptions* /*options*/,
                                        IFileSystemCallback* /*callback*/,
                                        void* /*cookie*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* /*sourcePaths*/,
                                        unsigned long /*count*/,
                                        const wchar_t* /*destinationFolder*/,
                                        FileSystemFlags /*flags*/,
                                        const FileSystemOptions* /*options*/,
                                        IFileSystemCallback* /*callback*/,
                                        void* /*cookie*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* /*paths*/,
                                          unsigned long /*count*/,
                                          FileSystemFlags /*flags*/,
                                          const FileSystemOptions* /*options*/,
                                          IFileSystemCallback* /*callback*/,
                                          void* /*cookie*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* /*items*/,
                                          unsigned long /*count*/,
                                          FileSystemFlags /*flags*/,
                                          const FileSystemOptions* /*options*/,
                                          IFileSystemCallback* /*callback*/,
                                          void* /*cookie*/) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        if (! jsonUtf8)
        {
            return E_POINTER;
        }

        *jsonUtf8 = nullptr;
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

private:
    inline static const PluginMetaData kMetaData{
        L"builtin/file-system",
        L"file",
        L"File System",
        L"Test built-in file system stub",
        L"ViewerSqliteTests",
        L"1",
    };

    std::atomic_ulong _refCount{1};
};

[[nodiscard]] TempDatabase CreateDatabase(std::wstring& errorText)
{
    TempDatabase tempDb;

    const auto uniquePart = std::format(L"viewer-sqlite-tests-{}-{}.sqlite", GetCurrentProcessId(), GetTickCount64());
    tempDb.path           = std::filesystem::temp_directory_path() / uniquePart;

    sqlite3* raw     = nullptr;
    const int openRc = sqlite3_open16(tempDb.path.c_str(), &raw);
    unique_sqlite3 db(raw, sqlite3_close_v2);
    if (openRc != SQLITE_OK || ! db)
    {
        errorText = L"Failed to create the temporary SQLite database.";
        return {};
    }

    const std::vector<const char*> statements = {"CREATE TABLE bigdata (id INTEGER PRIMARY KEY, name TEXT NOT NULL, payload TEXT);",
                                                 "CREATE VIEW bigdata_view AS SELECT id, name FROM bigdata;"};

    for (const char* statement : statements)
    {
        if (! Exec(db.get(), statement, errorText))
        {
            return {};
        }
    }

    if (! Exec(db.get(), "BEGIN TRANSACTION;", errorText))
    {
        return {};
    }

    for (int index = 1; index <= 750; ++index)
    {
        const std::string sql = std::format("INSERT INTO bigdata (id, name, payload) VALUES ({}, 'row-{}', 'payload-{}');", index, index, index);
        if (! Exec(db.get(), sql.c_str(), errorText))
        {
            return {};
        }
    }

    if (! Exec(db.get(), "COMMIT;", errorText))
    {
        return {};
    }

    return tempDb;
}

[[nodiscard]] bool ContainsTable(const std::vector<ViewerSqliteEngine::TableInfo>& tables, std::wstring_view name, std::wstring_view kind) noexcept
{
    return std::any_of(tables.begin(), tables.end(), [&](const ViewerSqliteEngine::TableInfo& table) { return table.name == name && table.kind == kind; });
}

bool Check(bool condition, std::wstring_view message, bool& success)
{
    if (condition)
    {
        std::wcout << std::format(L"[PASS] {}\n", message) << std::flush;
        return true;
    }

    std::wcout << std::format(L"[FAIL] {}\n", message) << std::flush;
    success = false;
    return false;
}

[[nodiscard]] bool TestListTables(const ViewerSqliteEngine::DatabaseSource& source)
{
    bool success = true;

    std::vector<ViewerSqliteEngine::TableInfo> tables;
    std::wstring errorText;
    const HRESULT hr = source.ListTables(tables, errorText);

    Check(SUCCEEDED(hr), L"table enumeration succeeds", success);
    Check(errorText.empty(), L"table enumeration does not return an error", success);
    Check(ContainsTable(tables, L"bigdata", L"table"), L"bigdata table is present", success);
    Check(ContainsTable(tables, L"bigdata_view", L"view"), L"bigdata_view is present", success);

    return success;
}

[[nodiscard]] bool TestPagedReads(const ViewerSqliteEngine::DatabaseSource& source)
{
    bool success = true;

    const auto middlePage = source.LoadTablePage(L"bigdata", 200, 400);
    Check(SUCCEEDED(middlePage.hr), L"middle page load succeeds", success);
    Check(middlePage.page.rows.size() == 200u, L"middle page contains exactly 200 rows", success);
    Check(middlePage.page.hasMore, L"middle page reports more rows", success);
    Check(! middlePage.page.rows.empty() && ! middlePage.page.rows.front().empty() && middlePage.page.rows.front()[0] == L"401",
          L"middle page starts at row id 401",
          success);

    const auto lastPage = source.LoadTablePage(L"bigdata", 200, 600);
    Check(SUCCEEDED(lastPage.hr), L"last page load succeeds", success);
    Check(lastPage.page.rows.size() == 150u, L"last page contains remaining 150 rows", success);
    Check(! lastPage.page.hasMore, L"last page does not report extra rows", success);
    Check(! lastPage.page.rows.empty() && ! lastPage.page.rows.back().empty() && lastPage.page.rows.back()[0] == L"750",
          L"last page ends at row id 750",
          success);

    return success;
}

[[nodiscard]] bool TestSortedPagedReads(const ViewerSqliteEngine::DatabaseSource& source)
{
    bool success = true;

    const auto descendingPage = source.LoadTablePage(L"bigdata", 25, 0, 0, ViewerSqliteEngine::TableSortDirection::Descending);
    Check(SUCCEEDED(descendingPage.hr), L"descending table preview succeeds", success);
    Check(! descendingPage.page.rows.empty() && ! descendingPage.page.rows.front().empty() && descendingPage.page.rows.front()[0] == L"750",
          L"descending table preview starts at row id 750",
          success);
    Check(descendingPage.page.executedSql.find(L"ORDER BY 1 DESC") != std::wstring::npos, L"descending table preview emits source-side ORDER BY", success);

    const auto ascendingPage = source.LoadTablePage(L"bigdata", 10, 10, 0, ViewerSqliteEngine::TableSortDirection::Ascending);
    Check(SUCCEEDED(ascendingPage.hr), L"ascending sorted page with offset succeeds", success);
    Check(! ascendingPage.page.rows.empty() && ! ascendingPage.page.rows.front().empty() && ascendingPage.page.rows.front()[0] == L"11",
          L"ascending sorted page offset is preserved",
          success);

    return success;
}

[[nodiscard]] bool TestReadOnlyQueries(const ViewerSqliteEngine::DatabaseSource& source)
{
    bool success = true;

    const auto validation = source.ValidateReadOnlyQuery(L"SELECT id, name FROM bigdata WHERE id <= 3 ORDER BY id DESC;");
    Check(SUCCEEDED(validation.hr) && validation.accepted, L"read-only SELECT query is accepted", success);

    const auto result = source.ExecuteReadOnlyQuery(L"SELECT id, name FROM bigdata WHERE id <= 3 ORDER BY id DESC;", 10);
    Check(SUCCEEDED(result.hr), L"read-only SELECT executes successfully", success);
    Check(result.page.rows.size() == 3u, L"read-only SELECT returns three rows", success);
    Check(result.page.columns.size() == 2u, L"read-only SELECT returns two columns", success);
    Check(! result.page.rows.empty() && result.page.rows.front()[0] == L"3", L"read-only SELECT preserves ordering", success);

    const auto rejectedWrite = source.ValidateReadOnlyQuery(L"UPDATE bigdata SET name = 'changed' WHERE id = 1;");
    Check(FAILED(rejectedWrite.hr) && ! rejectedWrite.accepted, L"write query is rejected", success);

    const auto rejectedBatch = source.ValidateReadOnlyQuery(L"SELECT 1; SELECT 2;");
    Check(FAILED(rejectedBatch.hr) && ! rejectedBatch.accepted, L"multiple statements are rejected", success);

    return success;
}

[[nodiscard]] bool TestViewerWindowUsesDxUiHostWithNoVisibleChildControls(const std::filesystem::path& databasePath) noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"test executable path resolves", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path dllPath = std::filesystem::path(modulePath.data()).parent_path() / L"Plugins" / L"ViewerSqlite.dll";
    Check(std::filesystem::exists(dllPath), L"ViewerSqlite.dll is present for viewer-window validation", success);
    if (! std::filesystem::exists(dllPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(dllPath.parent_path().parent_path().c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(dllPath.parent_path().c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(dllPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerSqlite.dll loads successfully", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerSqlite factory export is available", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerSqlitePluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerSqlite factory creates an IViewer instance", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    ViewerClosedCounter closeCounter;
    const HRESULT setCallbackHr = viewer->SetCallback(&closeCounter, &closeCounter);
    Check(SUCCEEDED(setCallbackHr), L"viewer callback registration succeeds", success);
    if (FAILED(setCallbackHr))
    {
        return false;
    }

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerSqliteWindowClassName);

    BuiltinFileSystemStub fileSystem;
    const wchar_t* otherFiles[] = {databasePath.c_str()};
    ViewerOpenContext context{};
    context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName        = L"File System";
    context.focusedPath           = databasePath.c_str();
    context.otherFiles            = otherFiles;
    context.otherFileCount        = static_cast<unsigned long>(std::size(otherFiles));
    context.focusedOtherFileIndex = 0u;

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"viewer window open succeeds", success);
    if (FAILED(openHr))
    {
        return false;
    }

    HWND viewerWindow       = nullptr;
    const bool openedWindow = PumpUntil(
        [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerSqliteWindowClassName, existingWindows);
        return viewerWindow != nullptr;
    },
        5000ms);
    Check(openedWindow, L"viewer window becomes visible", success);
    if (! openedWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    Check(CountVisibleChildWindows(viewerWindow) == 0u, L"viewer window does not expose visible child controls", success);
    static_cast<void>(PumpUntil([]() noexcept { return false; }, 250ms));

    LRESULT wmGetObject       = 0;
    const bool gotWmGetObject = PumpUntil(
        [&]() noexcept
    {
        wmGetObject = SendMessageW(viewerWindow, WM_GETOBJECT, 0, static_cast<LPARAM>(UiaRootObjectId));
        return wmGetObject != 0;
    },
        5000ms);
    Check(gotWmGetObject, L"viewer window answers WM_GETOBJECT", success);

#if defined(_DEBUG) && ! defined(__SANITIZE_ADDRESS__)
    UiaViewerSubtreeStats uiaStats{};
    const bool uiaReady = PumpUntil(
        [&]() noexcept
    {
        if (! TryCollectVisibleUiaViewerSubtreeStats(viewerWindow, uiaStats))
        {
            return false;
        }

        return uiaStats.visibleElementCount >= 10u && uiaStats.comboBoxControlCount >= 2u && uiaStats.editControlCount >= 1u &&
               uiaStats.buttonControlCount >= 5u && uiaStats.dataGridControlCount >= 1u && uiaStats.dataItemControlCount >= 1u &&
               uiaStats.valuePatternCount >= 1u && uiaStats.selectionPatternCount >= 1u;
    },
        10000ms);
    Check(uiaReady, L"viewer window exposes a visible UIA subtree for the DX form and grid", success);
    Check(uiaStats.visibleElementCount >= 10u, L"viewer window exposes multiple visible UIA descendants for the DX surface", success);
    Check(uiaStats.comboBoxControlCount >= 2u, L"viewer window exposes visible UIA ComboBox descendants", success);
    Check(uiaStats.editControlCount >= 1u, L"viewer window exposes a visible UIA Edit descendant", success);
    Check(uiaStats.buttonControlCount >= 5u, L"viewer window exposes visible UIA Button descendants", success);
    Check(uiaStats.dataGridControlCount >= 1u, L"viewer window exposes a visible UIA DataGrid descendant", success);
    Check(uiaStats.dataItemControlCount >= 1u, L"viewer window exposes visible UIA DataItem row descendants", success);
    Check(uiaStats.valuePatternCount >= 1u, L"viewer window exposes live UIA ValuePattern support", success);
    Check(uiaStats.selectionPatternCount >= 1u, L"viewer window exposes live UIA SelectionPattern support on the result grid", success);

    WndMsg::ViewerSqliteDebugSnapshot snapshot{};
    Check(TryGetViewerSqliteDebugSnapshot(viewerWindow, snapshot), L"viewer window exposes a debug snapshot for row-selection validation", success);
    Check(snapshot.hasStatusStrip, L"viewer window exposes a retained DxUi StatusStrip", success);
    Check(snapshot.statusStripVisible, L"viewer window keeps the DxUi StatusStrip visible", success);
    Check(snapshot.statusStripHeightDip >= 20.0f,
          std::format(L"viewer window keeps the DxUi StatusStrip at a usable height ({:.2f} DIP)", snapshot.statusStripHeightDip),
          success);
    Check(snapshot.rowCount >= 2u, L"viewer window exposes at least two rows for row-selection validation", success);
    if (snapshot.rowCount >= 2u)
    {
        Check(DebugSelectViewerSqliteGridRow(viewerWindow, 0u), L"viewer window debug row-selection hook can select the first row", success);

        UiaSelectionState initialSelection{};
        const bool initialSelectionReady = PumpUntil(
            [&]() noexcept
        {
            return TryCollectVisibleUiaViewerGridSelectionState(viewerWindow, initialSelection) && initialSelection.selectedCount == 1u &&
                   initialSelection.selectedControlType == UIA_DataItemControlTypeId && initialSelection.selectedHasSelectionItemPattern &&
                   ! initialSelection.selectedName.empty();
        },
            5000ms);
        Check(initialSelectionReady, L"viewer window exposes an initial selected UIA grid row", success);

        if (initialSelectionReady)
        {
            Check(DebugSelectViewerSqliteGridRow(viewerWindow, 1u), L"viewer window debug row-selection hook can select the second row", success);

            UiaSelectionState changedSelection{};
            const bool changedSelectionReady = PumpUntil(
                [&]() noexcept
            {
                return TryCollectVisibleUiaViewerGridSelectionState(viewerWindow, changedSelection) && changedSelection.selectedCount == 1u &&
                       changedSelection.selectedControlType == UIA_DataItemControlTypeId && changedSelection.selectedHasSelectionItemPattern &&
                       ! changedSelection.selectedName.empty() && changedSelection.selectedName != initialSelection.selectedName;
            },
                5000ms);
            Check(changedSelectionReady, L"viewer window UIA grid selection tracks a real row change", success);
        }
    }
#endif

    static_cast<void>(viewer->SetCallback(nullptr, nullptr));
    Check(closeCounter.closedCount.load(std::memory_order_relaxed) == 0u,
          L"clearing the viewer callback before close does not deliver ViewerClosed immediately",
          success);

    static_cast<void>(PumpUntil([]() noexcept { return false; }, 250ms));
    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"viewer window close succeeds", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), L"viewer window closes cleanly", success);
    Check(closeCounter.closedCount.load(std::memory_order_relaxed) == 0u,
          L"viewer does not invoke ViewerClosed after SetCallback(nullptr, nullptr) returns",
          success);

    viewer.reset();
    Check(true, L"viewer COM instance releases cleanly", success);
    pluginModule.reset();
    Check(true, L"ViewerSqlite.dll unload returns", success);
    return success;
}

#ifdef _DEBUG
[[nodiscard]] bool TestViewerWindowLongRunScrollingStaysBounded(const std::filesystem::path& databasePath) noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"test executable path resolves", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path dllPath = std::filesystem::path(modulePath.data()).parent_path() / L"Plugins" / L"ViewerSqlite.dll";
    Check(std::filesystem::exists(dllPath), L"ViewerSqlite.dll is present for long-run scrolling validation", success);
    if (! std::filesystem::exists(dllPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(dllPath.parent_path().parent_path().c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(dllPath.parent_path().c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(dllPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerSqlite.dll loads successfully for long-run scrolling validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerSqlite factory export is available for long-run scrolling validation", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerSqlitePluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerSqlite factory creates an IViewer instance for long-run scrolling validation", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerSqliteWindowClassName);

    BuiltinFileSystemStub fileSystem;
    const wchar_t* otherFiles[] = {databasePath.c_str()};
    ViewerOpenContext context{};
    context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName        = L"File System";
    context.focusedPath           = databasePath.c_str();
    context.otherFiles            = otherFiles;
    context.otherFileCount        = static_cast<unsigned long>(std::size(otherFiles));
    context.focusedOtherFileIndex = 0u;

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"viewer window open succeeds for long-run scrolling validation", success);
    if (FAILED(openHr))
    {
        return false;
    }

    HWND viewerWindow       = nullptr;
    const bool openedWindow = PumpUntil(
        [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerSqliteWindowClassName, existingWindows);
        return viewerWindow != nullptr;
    },
        5000ms);
    Check(openedWindow, L"viewer window becomes visible for long-run scrolling validation", success);
    if (! openedWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    LRESULT wmGetObject       = 0;
    const bool gotWmGetObject = PumpUntil(
        [&]() noexcept
    {
        wmGetObject = SendMessageW(viewerWindow, WM_GETOBJECT, 0, static_cast<LPARAM>(UiaRootObjectId));
        return wmGetObject != 0;
    },
        5000ms);
    Check(gotWmGetObject, L"viewer window answers WM_GETOBJECT for long-run scrolling validation", success);

    UiaViewerSubtreeStats uiaStats{};
    const bool uiaReady = PumpUntil(
        [&]() noexcept
    {
        if (! TryCollectVisibleUiaViewerSubtreeStats(viewerWindow, uiaStats))
        {
            return false;
        }

        return uiaStats.dataGridControlCount >= 1u && uiaStats.dataItemControlCount >= 1u && uiaStats.selectionPatternCount >= 1u;
    },
        10000ms);
    Check(uiaReady, L"viewer window exposes a scrollable UIA grid subtree for long-run scrolling validation", success);
    if (! uiaReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    WndMsg::ViewerSqliteDebugSnapshot baselineSnapshot{};
    const bool baselineReady = PumpUntil(
        [&]() noexcept
    {
        if (! TryGetViewerSqliteDebugSnapshot(viewerWindow, baselineSnapshot))
        {
            return false;
        }

        return baselineSnapshot.pendingAsyncWork == 0u && baselineSnapshot.rowCount > baselineSnapshot.visibleRowCount &&
               baselineSnapshot.visibleRowCount > 0u && baselineSnapshot.visibleColumnCount > 0u && baselineSnapshot.visibleCellCount > 0u &&
               baselineSnapshot.renderCount > 0u && baselineSnapshot.hasVerticalScrollbar;
    },
        10000ms);
    Check(baselineReady, L"viewer window reaches an idle populated DX grid state", success);
    if (! baselineReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    const size_t initialRowCount           = baselineSnapshot.rowCount;
    const size_t initialVisibleRowCount    = baselineSnapshot.visibleRowCount;
    const size_t initialVisibleColumnCount = baselineSnapshot.visibleColumnCount;
    const uint64_t initialResizeCount      = baselineSnapshot.resizeCount;

    constexpr int kWheelDetentsPerChunk = -4;
    for (size_t chunkIndex = 0; chunkIndex < 8u; ++chunkIndex)
    {
        const uint64_t previousRenderCount = baselineSnapshot.renderCount;
        Check(DebugScrollViewerSqliteGridByWheelDetents(viewerWindow, kWheelDetentsPerChunk),
              std::format(L"viewer window scroll chunk {} dispatch succeeds", chunkIndex + 1u),
              success);

        WndMsg::ViewerSqliteDebugSnapshot chunkSnapshot{};
        const bool chunkReady = PumpUntil(
            [&]() noexcept
        {
            if (! TryGetViewerSqliteDebugSnapshot(viewerWindow, chunkSnapshot))
            {
                return false;
            }

            return chunkSnapshot.pendingAsyncWork == 0u && chunkSnapshot.renderCount > previousRenderCount;
        },
            5000ms);
        Check(chunkReady, std::format(L"viewer window scroll chunk {} renders a settled frame", chunkIndex + 1u), success);
        if (! chunkReady)
        {
            success = false;
            break;
        }

        Check(chunkSnapshot.rowCount == initialRowCount, std::format(L"viewer window scroll chunk {} keeps row count stable", chunkIndex + 1u), success);
        Check(chunkSnapshot.visibleRowCount > 0u && chunkSnapshot.visibleRowCount <= initialVisibleRowCount + 1u,
              std::format(L"viewer window scroll chunk {} keeps visible rows bounded", chunkIndex + 1u),
              success);
        Check(chunkSnapshot.visibleColumnCount == initialVisibleColumnCount,
              std::format(L"viewer window scroll chunk {} keeps visible columns stable", chunkIndex + 1u),
              success);
        Check(chunkSnapshot.visibleCellCount == chunkSnapshot.visibleRowCount * chunkSnapshot.visibleColumnCount,
              std::format(L"viewer window scroll chunk {} keeps visible cell accounting consistent", chunkIndex + 1u),
              success);
        Check(chunkSnapshot.hasVerticalScrollbar, std::format(L"viewer window scroll chunk {} keeps the vertical scrollbar visible", chunkIndex + 1u), success);
        Check(chunkSnapshot.resizeCount == initialResizeCount,
              std::format(L"viewer window scroll chunk {} avoids DX host resize churn", chunkIndex + 1u),
              success);
        Check(chunkSnapshot.resizeFailureCount == 0u, std::format(L"viewer window scroll chunk {} avoids DX host resize failures", chunkIndex + 1u), success);

        baselineSnapshot = chunkSnapshot;
    }

    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"viewer window close succeeds after long-run scrolling validation", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms),
          L"viewer window closes cleanly after long-run scrolling validation",
          success);

    static_cast<void>(viewer->SetCallback(nullptr, nullptr));
    viewer.reset();
    pluginModule.reset();
    return success;
}

[[nodiscard]] bool TestViewerWindowLongRunOpenCloseStaysStable(const std::filesystem::path& databasePath) noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"test executable path resolves for open-close churn validation", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path dllPath = std::filesystem::path(modulePath.data()).parent_path() / L"Plugins" / L"ViewerSqlite.dll";
    Check(std::filesystem::exists(dllPath), L"ViewerSqlite.dll is present for open-close churn validation", success);
    if (! std::filesystem::exists(dllPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(dllPath.parent_path().parent_path().c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(dllPath.parent_path().c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(dllPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerSqlite.dll loads successfully for open-close churn validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerSqlite factory export is available for open-close churn validation", success);
    if (! createFn)
    {
        return false;
    }

    BuiltinFileSystemStub fileSystem;
    const wchar_t* otherFiles[] = {databasePath.c_str()};

    for (size_t cycleIndex = 0; cycleIndex < 12u; ++cycleIndex)
    {
        wil::com_ptr<IViewer> viewer;
        const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
        const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerSqlitePluginId, viewer.put_void());
        Check(SUCCEEDED(createHr) && viewer != nullptr, std::format(L"viewer window cycle {} creates an IViewer instance", cycleIndex + 1u), success);
        if (FAILED(createHr) || ! viewer)
        {
            return false;
        }

        HWND viewerWindow  = nullptr;
        auto cleanupViewer = wil::scope_exit([&]() noexcept
        {
            if (viewer)
            {
                static_cast<void>(viewer->SetCallback(nullptr, nullptr));
                if (viewerWindow != nullptr && IsWindow(viewerWindow) != FALSE)
                {
                    static_cast<void>(viewer->Close());
                    static_cast<void>(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms));
                }
                viewer.reset();
            }
        });

        const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerSqliteWindowClassName);

        ViewerOpenContext context{};
        context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
        context.fileSystemName        = L"File System";
        context.focusedPath           = databasePath.c_str();
        context.otherFiles            = otherFiles;
        context.otherFileCount        = static_cast<unsigned long>(std::size(otherFiles));
        context.focusedOtherFileIndex = 0u;

        const HRESULT openHr = viewer->Open(&context);
        Check(SUCCEEDED(openHr), std::format(L"viewer window cycle {} opens successfully", cycleIndex + 1u), success);
        if (FAILED(openHr))
        {
            success = false;
            break;
        }

        const bool openedWindow = PumpUntil(
            [&]() noexcept
        {
            viewerWindow = FindNewVisibleWindowByClass(kViewerSqliteWindowClassName, existingWindows);
            return viewerWindow != nullptr;
        },
            5000ms);
        Check(openedWindow, std::format(L"viewer window cycle {} becomes visible", cycleIndex + 1u), success);
        if (! openedWindow || ! viewerWindow)
        {
            success = false;
            break;
        }

        Check(CountVisibleChildWindows(viewerWindow) == 0u,
              std::format(L"viewer window cycle {} keeps visible child fallback at zero", cycleIndex + 1u),
              success);
        static_cast<void>(PumpUntil([]() noexcept { return false; }, 250ms));

        LRESULT wmGetObject       = 0;
        const bool gotWmGetObject = PumpUntil(
            [&]() noexcept
        {
            wmGetObject = SendMessageW(viewerWindow, WM_GETOBJECT, 0, static_cast<LPARAM>(UiaRootObjectId));
            return wmGetObject != 0;
        },
            5000ms);
        Check(gotWmGetObject, std::format(L"viewer window cycle {} answers WM_GETOBJECT", cycleIndex + 1u), success);

        UiaViewerSubtreeStats uiaStats{};
        const bool uiaReady = PumpUntil(
            [&]() noexcept
        {
            if (! TryCollectVisibleUiaViewerSubtreeStats(viewerWindow, uiaStats))
            {
                return false;
            }

            return uiaStats.visibleElementCount >= 10u && uiaStats.comboBoxControlCount >= 2u && uiaStats.editControlCount >= 1u &&
                   uiaStats.buttonControlCount >= 5u && uiaStats.dataGridControlCount >= 1u && uiaStats.dataItemControlCount >= 1u &&
                   uiaStats.valuePatternCount >= 1u && uiaStats.selectionPatternCount >= 1u;
        },
            10000ms);
        Check(uiaReady, std::format(L"viewer window cycle {} exposes the visible DX UIA subtree on reopen", cycleIndex + 1u), success);

        WndMsg::ViewerSqliteDebugSnapshot snapshot{};
        const bool snapshotReady = WaitForViewerSnapshot(viewerWindow,
                                                         [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
        {
            return value.pendingAsyncWork == 0u && value.rowCount > 0u && value.visibleRowCount > 0u && value.visibleColumnCount > 0u &&
                   value.visibleCellCount > 0u && value.renderCount > 0u;
        },
                                                         10000ms,
                                                         &snapshot);
        Check(snapshotReady, std::format(L"viewer window cycle {} reaches an idle populated DX grid state", cycleIndex + 1u), success);
        Check(snapshot.rowCount > 0u, std::format(L"viewer window cycle {} keeps result rows populated after reopen", cycleIndex + 1u), success);

        const HRESULT closeHr = viewer->Close();
        Check(SUCCEEDED(closeHr), std::format(L"viewer window cycle {} closes successfully", cycleIndex + 1u), success);
        const bool closedWindow = SUCCEEDED(closeHr) && PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms);
        Check(closedWindow, std::format(L"viewer window cycle {} closes without leaving a lingering window handle", cycleIndex + 1u), success);

        static_cast<void>(viewer->SetCallback(nullptr, nullptr));
        viewer.reset();
        cleanupViewer.release();
    }

    pluginModule.reset();
    return success;
}

[[nodiscard]] bool TestViewerWindowPagingAndSortFlows(const std::filesystem::path& databasePath) noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"test executable path resolves for paging and sort validation", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path dllPath = std::filesystem::path(modulePath.data()).parent_path() / L"Plugins" / L"ViewerSqlite.dll";
    Check(std::filesystem::exists(dllPath), L"ViewerSqlite.dll is present for paging and sort validation", success);
    if (! std::filesystem::exists(dllPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(dllPath.parent_path().parent_path().c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(dllPath.parent_path().c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(dllPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerSqlite.dll loads successfully for paging and sort validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerSqlite factory export is available for paging and sort validation", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerSqlitePluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerSqlite factory creates an IViewer instance for paging and sort validation", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerSqliteWindowClassName);

    BuiltinFileSystemStub fileSystem;
    const wchar_t* otherFiles[] = {databasePath.c_str()};
    ViewerOpenContext context{};
    context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName        = L"File System";
    context.focusedPath           = databasePath.c_str();
    context.otherFiles            = otherFiles;
    context.otherFileCount        = static_cast<unsigned long>(std::size(otherFiles));
    context.focusedOtherFileIndex = 0u;

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"viewer window open succeeds for paging and sort validation", success);
    if (FAILED(openHr))
    {
        return false;
    }

    HWND viewerWindow       = nullptr;
    const bool openedWindow = PumpUntil(
        [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerSqliteWindowClassName, existingWindows);
        return viewerWindow != nullptr;
    },
        5000ms);
    Check(openedWindow, L"viewer window becomes visible for paging and sort validation", success);
    if (! openedWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    WndMsg::ViewerSqliteDebugSnapshot snapshot{};
    const bool initialReady = WaitForViewerSnapshot(viewerWindow,
                                                    [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.tablePreviewMode && value.rowCount == 200u && value.rowOffset == 0u && value.firstRowPrimaryKey == 1u &&
               value.lastRowPrimaryKey == 200u && ! value.prevButtonEnabled && value.nextButtonEnabled && value.sortDirection == 0u;
    },
                                                    10000ms,
                                                    &snapshot);
    Check(initialReady, L"viewer window reaches the initial unsorted table-preview page", success);
    if (! initialReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    Check(DebugInvokeViewerSqlitePageCommand(viewerWindow, WndMsg::ViewerSqliteDebugPageCommand::Next),
          L"viewer paging debug hook can request the next table page",
          success);

    const bool nextPageReady = WaitForViewerSnapshot(viewerWindow,
                                                     [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.rowOffset == 200u && value.rowCount == 200u && value.firstRowPrimaryKey == 201u &&
               value.lastRowPrimaryKey == 400u && value.prevButtonEnabled && value.nextButtonEnabled && value.sortDirection == 0u;
    },
                                                     10000ms,
                                                     &snapshot);
    Check(nextPageReady, L"viewer next-page navigation loads the second table page", success);

    Check(DebugInvokeViewerSqlitePageCommand(viewerWindow, WndMsg::ViewerSqliteDebugPageCommand::Previous),
          L"viewer paging debug hook can request the previous table page",
          success);

    const bool previousPageReady = WaitForViewerSnapshot(viewerWindow,
                                                         [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.rowOffset == 0u && value.rowCount == 200u && value.firstRowPrimaryKey == 1u &&
               value.lastRowPrimaryKey == 200u && ! value.prevButtonEnabled && value.nextButtonEnabled && value.sortDirection == 0u;
    },
                                                         10000ms,
                                                         &snapshot);
    Check(previousPageReady, L"viewer previous-page navigation returns to the first table page", success);

    Check(DebugCycleViewerSqliteSortColumn(viewerWindow, 0u), L"viewer sort debug hook can request ascending sort on the first column", success);

    const bool ascendingSortReady = WaitForViewerSnapshot(viewerWindow,
                                                          [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.rowOffset == 0u && value.sortColumnIndex == 0u && value.sortDirection == 1u &&
               value.firstRowPrimaryKey == 1u && value.lastRowPrimaryKey == 200u && ! value.prevButtonEnabled && value.nextButtonEnabled;
    },
                                                          10000ms,
                                                          &snapshot);
    Check(ascendingSortReady, L"viewer ascending sort reloads the first page with sort metadata", success);

    Check(DebugCycleViewerSqliteSortColumn(viewerWindow, 0u), L"viewer sort debug hook can request descending sort on the first column", success);

    const bool descendingSortReady = WaitForViewerSnapshot(viewerWindow,
                                                           [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.rowOffset == 0u && value.sortColumnIndex == 0u && value.sortDirection == 2u &&
               value.firstRowPrimaryKey == 750u && value.lastRowPrimaryKey == 551u && ! value.prevButtonEnabled && value.nextButtonEnabled;
    },
                                                           10000ms,
                                                           &snapshot);
    Check(descendingSortReady, L"viewer descending sort reloads the first page in reverse key order", success);

    Check(DebugInvokeViewerSqlitePageCommand(viewerWindow, WndMsg::ViewerSqliteDebugPageCommand::Next),
          L"viewer paging remains active after descending sort",
          success);

    const bool descendingNextPageReady = WaitForViewerSnapshot(viewerWindow,
                                                               [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.rowOffset == 200u && value.sortColumnIndex == 0u && value.sortDirection == 2u &&
               value.firstRowPrimaryKey == 550u && value.lastRowPrimaryKey == 351u && value.prevButtonEnabled && value.nextButtonEnabled;
    },
                                                               10000ms,
                                                               &snapshot);
    Check(descendingNextPageReady, L"viewer descending sort preserves paged navigation across subsequent pages", success);

    Check(DebugCycleViewerSqliteSortColumn(viewerWindow, 0u), L"viewer sort debug hook can cycle the first column back to unsorted", success);

    const bool unsortedResetReady = WaitForViewerSnapshot(viewerWindow,
                                                          [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.rowOffset == 0u && value.sortColumnIndex == static_cast<size_t>(-1) && value.sortDirection == 0u &&
               value.firstRowPrimaryKey == 1u && value.lastRowPrimaryKey == 200u && ! value.prevButtonEnabled && value.nextButtonEnabled;
    },
                                                          10000ms,
                                                          &snapshot);
    Check(unsortedResetReady, L"viewer cycling sort off resets back to the unsorted first page", success);

    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"viewer window close succeeds after paging and sort validation", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms),
          L"viewer window closes cleanly after paging and sort validation",
          success);

    static_cast<void>(viewer->SetCallback(nullptr, nullptr));
    viewer.reset();
    pluginModule.reset();
    return success;
}

#if ! defined(__SANITIZE_ADDRESS__)
[[nodiscard]] bool TestViewerWindowSelectionPatternSurvivesPagingAndSortFlows(const std::filesystem::path& databasePath) noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"test executable path resolves for viewer UIA selection churn validation", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path dllPath = std::filesystem::path(modulePath.data()).parent_path() / L"Plugins" / L"ViewerSqlite.dll";
    Check(std::filesystem::exists(dllPath), L"ViewerSqlite.dll is present for viewer UIA selection churn validation", success);
    if (! std::filesystem::exists(dllPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(dllPath.parent_path().parent_path().c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(dllPath.parent_path().c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(dllPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerSqlite.dll loads successfully for viewer UIA selection churn validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerSqlite factory export is available for viewer UIA selection churn validation", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerSqlitePluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerSqlite factory creates an IViewer instance for viewer UIA selection churn validation", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerSqliteWindowClassName);

    BuiltinFileSystemStub fileSystem;
    const wchar_t* otherFiles[] = {databasePath.c_str()};
    ViewerOpenContext context{};
    context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName        = L"File System";
    context.focusedPath           = databasePath.c_str();
    context.otherFiles            = otherFiles;
    context.otherFileCount        = static_cast<unsigned long>(std::size(otherFiles));
    context.focusedOtherFileIndex = 0u;

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"viewer window open succeeds for viewer UIA selection churn validation", success);
    if (FAILED(openHr))
    {
        return false;
    }

    HWND viewerWindow       = nullptr;
    const bool openedWindow = PumpUntil(
        [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerSqliteWindowClassName, existingWindows);
        return viewerWindow != nullptr;
    },
        5000ms);
    Check(openedWindow, L"viewer window becomes visible for viewer UIA selection churn validation", success);
    if (! openedWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    WndMsg::ViewerSqliteDebugSnapshot snapshot{};
    const bool initialReady = WaitForViewerSnapshot(viewerWindow,
                                                    [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.tablePreviewMode && value.rowCount == 200u && value.rowOffset == 0u && value.firstRowPrimaryKey == 1u &&
               value.lastRowPrimaryKey == 200u && ! value.prevButtonEnabled && value.nextButtonEnabled && value.sortDirection == 0u;
    },
                                                    10000ms,
                                                    &snapshot);
    Check(initialReady, L"viewer window reaches the initial unsorted table-preview page for viewer UIA selection churn validation", success);
    if (! initialReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    const auto selectAndValidateUiaRow = [&](size_t rowIndex, uint64_t expectedPrimaryKey, std::wstring_view phaseLabel, std::wstring* selectedNameOut) noexcept
    {
        Check(DebugSelectViewerSqliteGridRow(viewerWindow, rowIndex), std::format(L"{} can select the requested visible result row", phaseLabel), success);

        const bool snapshotReady = WaitForViewerSnapshot(viewerWindow,
                                                         [&](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
        {
            return value.pendingAsyncWork == 0u && value.selectionCount == 1u &&
                   (expectedPrimaryKey == 0u || value.primarySelectedRowId == expectedPrimaryKey) && value.visibleRowCount > 0u &&
                   value.visibleColumnCount > 0u && value.visibleCellCount == value.visibleRowCount * value.visibleColumnCount;
        },
                                                         5000ms,
                                                         &snapshot);
        Check(snapshotReady,
              std::format(
                  L"{} keeps the debug snapshot selection anchored to {}", phaseLabel, expectedPrimaryKey == 0u ? L"a real selected row" : L"the expected row"),
              success);

        UiaSelectionState selectionState{};
        const bool uiaReady = PumpUntil(
            [&]() noexcept
        {
            return TryCollectVisibleUiaViewerGridSelectionState(viewerWindow, selectionState) && selectionState.selectedCount == 1u &&
                   selectionState.selectedControlType == UIA_DataItemControlTypeId && selectionState.selectedHasSelectionItemPattern &&
                   selectionState.selectedVisible && ! selectionState.selectedName.empty();
        },
            5000ms);
        Check(uiaReady, std::format(L"{} keeps a visible selected UIA DataItem with SelectionItemPattern", phaseLabel), success);

        if (uiaReady && selectedNameOut)
        {
            *selectedNameOut = selectionState.selectedName;
        }

        return snapshotReady && uiaReady;
    };

    std::wstring firstPageSelectionName;
    const bool firstSelectionReady = selectAndValidateUiaRow(1u, 2u, L"initial unsorted page selection", &firstPageSelectionName);
    if (! firstSelectionReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    Check(DebugInvokeViewerSqlitePageCommand(viewerWindow, WndMsg::ViewerSqliteDebugPageCommand::Next),
          L"viewer paging debug hook can request the next page for viewer UIA selection churn validation",
          success);

    const bool nextPageReady = WaitForViewerSnapshot(viewerWindow,
                                                     [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.rowOffset == 200u && value.rowCount == 200u && value.firstRowPrimaryKey == 201u &&
               value.lastRowPrimaryKey == 400u && value.prevButtonEnabled && value.nextButtonEnabled && value.sortDirection == 0u;
    },
                                                     10000ms,
                                                     &snapshot);
    Check(nextPageReady, L"viewer next-page navigation settles before viewer UIA selection churn validation", success);
    if (! nextPageReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    std::wstring secondPageSelectionName;
    const bool secondSelectionReady = selectAndValidateUiaRow(1u, 202u, L"next-page selection", &secondPageSelectionName);
    Check(secondSelectionReady && secondPageSelectionName != firstPageSelectionName,
          L"viewer UIA selected row name changes after paging to a different result row",
          success);
    if (! secondSelectionReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    Check(DebugCycleViewerSqliteSortColumn(viewerWindow, 0u),
          L"viewer sort debug hook can request ascending sort before descending selection validation",
          success);
    const bool ascendingSortReady = WaitForViewerSnapshot(viewerWindow,
                                                          [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.rowOffset == 0u && value.sortColumnIndex == 0u && value.sortDirection == 1u &&
               value.firstRowPrimaryKey == 1u && value.lastRowPrimaryKey == 200u;
    },
                                                          10000ms,
                                                          &snapshot);
    Check(ascendingSortReady, L"viewer ascending sort settles before descending selection validation", success);
    if (! ascendingSortReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    Check(DebugCycleViewerSqliteSortColumn(viewerWindow, 0u),
          L"viewer sort debug hook can request descending sort for viewer UIA selection churn validation",
          success);
    const bool descendingSortReady = WaitForViewerSnapshot(viewerWindow,
                                                           [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.rowOffset == 0u && value.sortColumnIndex == 0u && value.sortDirection == 2u &&
               value.firstRowPrimaryKey == 750u && value.lastRowPrimaryKey == 551u && ! value.prevButtonEnabled && value.nextButtonEnabled;
    },
                                                           10000ms,
                                                           &snapshot);
    Check(descendingSortReady, L"viewer descending sort settles before viewer UIA selection churn validation", success);
    if (! descendingSortReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    std::wstring descendingSelectionName;
    const bool descendingSelectionReady = selectAndValidateUiaRow(1u, 0u, L"descending sorted page selection", &descendingSelectionName);
    if (! descendingSelectionReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    Check(DebugInvokeViewerSqlitePageCommand(viewerWindow, WndMsg::ViewerSqliteDebugPageCommand::Next),
          L"viewer paging stays active after descending sort for viewer UIA selection churn validation",
          success);
    const bool descendingNextPageReady = WaitForViewerSnapshot(viewerWindow,
                                                               [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.rowOffset == 200u && value.sortColumnIndex == 0u && value.sortDirection == 2u &&
               value.firstRowPrimaryKey == 550u && value.lastRowPrimaryKey == 351u && value.prevButtonEnabled && value.nextButtonEnabled;
    },
                                                               10000ms,
                                                               &snapshot);
    Check(descendingNextPageReady, L"viewer descending next page settles before final UIA selection validation", success);
    if (! descendingNextPageReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    std::wstring descendingNextPageSelectionName;
    const bool descendingNextSelectionReady = selectAndValidateUiaRow(1u, 0u, L"descending sorted next-page selection", &descendingNextPageSelectionName);
    Check(descendingNextSelectionReady, L"viewer UIA selection remains valid after descending paging churn", success);

    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"viewer window close succeeds after viewer UIA selection churn validation", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms),
          L"viewer window closes cleanly after viewer UIA selection churn validation",
          success);

    static_cast<void>(viewer->SetCallback(nullptr, nullptr));
    viewer.reset();
    pluginModule.reset();
    return success;
}
#endif

[[nodiscard]] bool TestViewerWindowTabTraversalMatchesExpectedOrder(const std::filesystem::path& databasePath) noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"test executable path resolves for tab traversal validation", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path dllPath = std::filesystem::path(modulePath.data()).parent_path() / L"Plugins" / L"ViewerSqlite.dll";
    Check(std::filesystem::exists(dllPath), L"ViewerSqlite.dll is present for tab traversal validation", success);
    if (! std::filesystem::exists(dllPath))
    {
        return false;
    }

    const std::filesystem::path alternatePath =
        databasePath.parent_path() / std::format(L"{}-alternate{}", databasePath.stem().wstring(), databasePath.extension().wstring());
    std::error_code copyEc;
    static_cast<void>(std::filesystem::copy_file(databasePath, alternatePath, std::filesystem::copy_options::overwrite_existing, copyEc));
    Check(! copyEc, L"alternate SQLite input is created for file-combo traversal coverage", success);
    if (copyEc)
    {
        return false;
    }
    const auto cleanupAlternate = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        static_cast<void>(std::filesystem::remove(alternatePath, cleanupEc));
    });

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(dllPath.parent_path().parent_path().c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(dllPath.parent_path().c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(dllPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerSqlite.dll loads successfully for tab traversal validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerSqlite factory export is available for tab traversal validation", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerSqlitePluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerSqlite factory creates an IViewer instance for tab traversal validation", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerSqliteWindowClassName);

    BuiltinFileSystemStub fileSystem;
    const std::wstring primaryPathText   = databasePath.wstring();
    const std::wstring alternatePathText = alternatePath.wstring();
    const wchar_t* otherFiles[]          = {primaryPathText.c_str(), alternatePathText.c_str()};
    ViewerOpenContext context{};
    context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName        = L"File System";
    context.focusedPath           = primaryPathText.c_str();
    context.otherFiles            = otherFiles;
    context.otherFileCount        = static_cast<unsigned long>(std::size(otherFiles));
    context.focusedOtherFileIndex = 0u;

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"viewer window open succeeds for tab traversal validation", success);
    if (FAILED(openHr))
    {
        return false;
    }

    HWND viewerWindow       = nullptr;
    const bool openedWindow = PumpUntil(
        [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerSqliteWindowClassName, existingWindows);
        return viewerWindow != nullptr;
    },
        5000ms);
    Check(openedWindow, L"viewer window becomes visible for tab traversal validation", success);
    if (! openedWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    WndMsg::ViewerSqliteDebugSnapshot snapshot{};
    const bool initialReady = WaitForViewerSnapshot(viewerWindow,
                                                    [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.tablePreviewMode && value.rowCount == 200u && value.rowOffset == 0u && value.visibleRowCount > 0u &&
               value.visibleColumnCount > 0u && value.visibleCellCount > 0u && ! value.prevButtonEnabled && value.nextButtonEnabled;
    },
                                                    10000ms,
                                                    &snapshot);
    Check(initialReady, L"viewer window reaches the initial idle table-preview state for tab traversal validation", success);
    if (! initialReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    Check(DebugInvokeViewerSqlitePageCommand(viewerWindow, WndMsg::ViewerSqliteDebugPageCommand::Next),
          L"viewer paging debug hook can move to the second page for tab traversal validation",
          success);

    const bool pageTwoReady = WaitForViewerSnapshot(viewerWindow,
                                                    [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.rowOffset == 200u && value.rowCount == 200u && value.firstRowPrimaryKey == 201u &&
               value.lastRowPrimaryKey == 400u && value.prevButtonEnabled && value.nextButtonEnabled && value.visibleRowCount > 0u &&
               value.visibleColumnCount > 0u && value.visibleCellCount > 0u;
    },
                                                    10000ms,
                                                    &snapshot);
    Check(pageTwoReady, L"viewer second page is ready before tab traversal validation", success);
    if (! pageTwoReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    Check(DebugSelectViewerSqliteGridRow(viewerWindow, 1u),
          L"viewer debug row-selection hook can select a stable second-row baseline for tab traversal validation",
          success);

    const bool selectionReady = WaitForViewerSnapshot(viewerWindow,
                                                      [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.rowOffset == 200u && value.selectionCount >= 1u && value.primarySelectedRowId == 202u &&
               value.prevButtonEnabled && value.nextButtonEnabled;
    },
                                                      10000ms,
                                                      &snapshot);
    Check(selectionReady, L"viewer grid selection settles before tab traversal validation", success);
    if (! selectionReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    Check(SetFocus(viewerWindow) != nullptr, L"viewer window accepts keyboard focus for tab traversal validation", success);
    const bool startingFocusReady = WaitForViewerSnapshot(viewerWindow, [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept {
        return value.pendingAsyncWork == 0u && value.rowOffset == 200u && value.primarySelectedRowId == 202u;
    }, 5000ms, &snapshot);
    Check(startingFocusReady, L"viewer snapshot is available after taking keyboard focus", success);
    if (! startingFocusReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    bool initialFocusReady = snapshot.focusTarget == WndMsg::ViewerSqliteDebugFocusTarget::ResultGrid;
    if (! initialFocusReady)
    {
        initialFocusReady = WaitForViewerSnapshot(viewerWindow,
                                                  [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
        {
            return value.pendingAsyncWork == 0u && value.focusTarget == WndMsg::ViewerSqliteDebugFocusTarget::ResultGrid && value.rowOffset == 200u &&
                   value.primarySelectedRowId == 202u;
        },
                                                  5000ms,
                                                  &snapshot);
    }
    Check(initialFocusReady, L"viewer activation leaves keyboard focus on the results grid", success);
    if (! initialFocusReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    const WndMsg::ViewerSqliteDebugSnapshot baseline = snapshot;
    const auto focusTargetName                       = [](WndMsg::ViewerSqliteDebugFocusTarget target) noexcept -> std::wstring_view
    {
        switch (target)
        {
            case WndMsg::ViewerSqliteDebugFocusTarget::None: return L"no focused target";
            case WndMsg::ViewerSqliteDebugFocusTarget::FileCombo: return L"file combo";
            case WndMsg::ViewerSqliteDebugFocusTarget::ReloadButton: return L"reload button";
            case WndMsg::ViewerSqliteDebugFocusTarget::TableCombo: return L"table combo";
            case WndMsg::ViewerSqliteDebugFocusTarget::PrevButton: return L"previous button";
            case WndMsg::ViewerSqliteDebugFocusTarget::NextButton: return L"next button";
            case WndMsg::ViewerSqliteDebugFocusTarget::QueryField: return L"query field";
            case WndMsg::ViewerSqliteDebugFocusTarget::RunButton: return L"run-query button";
            case WndMsg::ViewerSqliteDebugFocusTarget::TableButton: return L"table-preview button";
            case WndMsg::ViewerSqliteDebugFocusTarget::ResultGrid: return L"results grid";
            default: return L"unknown focus target";
        }
    };

    const auto focusAndStateStable = [&](const WndMsg::ViewerSqliteDebugSnapshot& value, WndMsg::ViewerSqliteDebugFocusTarget target) noexcept
    {
        return value.pendingAsyncWork == 0u && value.focusTarget == target && value.rowOffset == baseline.rowOffset && value.rowCount == baseline.rowCount &&
               value.selectionCount >= 1u && value.primarySelectedRowId == baseline.primarySelectedRowId && value.visibleRowCount == baseline.visibleRowCount &&
               value.visibleColumnCount == baseline.visibleColumnCount && value.visibleCellCount == baseline.visibleCellCount &&
               value.resizeFailureCount == baseline.resizeFailureCount && value.prevButtonEnabled == baseline.prevButtonEnabled &&
               value.nextButtonEnabled == baseline.nextButtonEnabled;
    };

    constexpr std::array<WndMsg::ViewerSqliteDebugFocusTarget, 9u> kForwardOrder{{
        WndMsg::ViewerSqliteDebugFocusTarget::FileCombo,
        WndMsg::ViewerSqliteDebugFocusTarget::ReloadButton,
        WndMsg::ViewerSqliteDebugFocusTarget::TableCombo,
        WndMsg::ViewerSqliteDebugFocusTarget::PrevButton,
        WndMsg::ViewerSqliteDebugFocusTarget::NextButton,
        WndMsg::ViewerSqliteDebugFocusTarget::QueryField,
        WndMsg::ViewerSqliteDebugFocusTarget::RunButton,
        WndMsg::ViewerSqliteDebugFocusTarget::TableButton,
        WndMsg::ViewerSqliteDebugFocusTarget::ResultGrid,
    }};

    for (const WndMsg::ViewerSqliteDebugFocusTarget target : kForwardOrder)
    {
        SendViewerSqliteTab(viewerWindow, false);
        const bool targetReady = WaitForViewerSnapshot(
            viewerWindow, [&](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept { return focusAndStateStable(value, target); }, 5000ms, &snapshot);
        Check(targetReady, std::format(L"forward Tab advances focus to the {}", focusTargetName(target)), success);
    }

    constexpr std::array<WndMsg::ViewerSqliteDebugFocusTarget, 9u> kReverseOrder{{
        WndMsg::ViewerSqliteDebugFocusTarget::TableButton,
        WndMsg::ViewerSqliteDebugFocusTarget::RunButton,
        WndMsg::ViewerSqliteDebugFocusTarget::QueryField,
        WndMsg::ViewerSqliteDebugFocusTarget::NextButton,
        WndMsg::ViewerSqliteDebugFocusTarget::PrevButton,
        WndMsg::ViewerSqliteDebugFocusTarget::TableCombo,
        WndMsg::ViewerSqliteDebugFocusTarget::ReloadButton,
        WndMsg::ViewerSqliteDebugFocusTarget::FileCombo,
        WndMsg::ViewerSqliteDebugFocusTarget::ResultGrid,
    }};

    for (const WndMsg::ViewerSqliteDebugFocusTarget target : kReverseOrder)
    {
        SendViewerSqliteTab(viewerWindow, true);
        const bool targetReady = WaitForViewerSnapshot(
            viewerWindow, [&](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept { return focusAndStateStable(value, target); }, 5000ms, &snapshot);
        Check(targetReady, std::format(L"reverse Tab advances focus to the {}", focusTargetName(target)), success);
    }

    SendViewerSqliteTab(viewerWindow, false);
    const bool fileComboFocused = WaitForViewerSnapshot(viewerWindow, [&](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept {
        return focusAndStateStable(value, WndMsg::ViewerSqliteDebugFocusTarget::FileCombo);
    }, 5000ms, &snapshot);
    Check(fileComboFocused, L"focus is on a viewer control before Escape focus-return validation", success);

    static_cast<void>(SendMessageW(viewerWindow, WM_KEYDOWN, VK_ESCAPE, 0));
    static_cast<void>(SendMessageW(viewerWindow, WM_KEYUP, VK_ESCAPE, 0));
    const bool escapeReturnedFocus = WaitForViewerSnapshot(viewerWindow, [&](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept {
        return focusAndStateStable(value, WndMsg::ViewerSqliteDebugFocusTarget::ResultGrid);
    }, 5000ms, &snapshot);
    Check(escapeReturnedFocus && IsWindow(viewerWindow) != FALSE, L"Escape from viewer controls returns focus to the results grid without closing", success);

    static_cast<void>(SendMessageW(viewerWindow, WM_KEYDOWN, VK_ESCAPE, 0));
    static_cast<void>(SendMessageW(viewerWindow, WM_KEYUP, VK_ESCAPE, 0));
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms),
          L"viewer window closes cleanly after Escape from the results grid",
          success);

    static_cast<void>(viewer->SetCallback(nullptr, nullptr));
    viewer.reset();
    pluginModule.reset();
    return success;
}

[[nodiscard]] bool TestViewerWindowThemeCycleKeepsGridLegible(const std::filesystem::path& databasePath) noexcept
{
    bool success = true;

    std::array<wchar_t, MAX_PATH + 1> modulePath{};
    const DWORD pathLength = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    Check(pathLength > 0 && pathLength < modulePath.size(), L"test executable path resolves for theme-cycle validation", success);
    if (pathLength == 0 || pathLength >= modulePath.size())
    {
        return false;
    }

    const std::filesystem::path dllPath = std::filesystem::path(modulePath.data()).parent_path() / L"Plugins" / L"ViewerSqlite.dll";
    Check(std::filesystem::exists(dllPath), L"ViewerSqlite.dll is present for theme-cycle validation", success);
    if (! std::filesystem::exists(dllPath))
    {
        return false;
    }

    static_cast<void>(SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_USER_DIRS));
    DLL_DIRECTORY_COOKIE buildCookie  = AddDllDirectory(dllPath.parent_path().parent_path().c_str());
    DLL_DIRECTORY_COOKIE pluginCookie = AddDllDirectory(dllPath.parent_path().c_str());
    auto removeDllDirectories         = wil::scope_exit([&]() noexcept
    {
        if (pluginCookie)
        {
            RemoveDllDirectory(pluginCookie);
        }
        if (buildCookie)
        {
            RemoveDllDirectory(buildCookie);
        }
    });

    wil::unique_hmodule pluginModule(LoadLibraryExW(dllPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS | LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR));
    Check(pluginModule.is_valid(), L"ViewerSqlite.dll loads successfully for theme-cycle validation", success);
    if (! pluginModule.is_valid())
    {
        return false;
    }

    const FARPROC createProc       = GetProcAddress(pluginModule.get(), "RedSalamanderCreate");
    RedSalamanderCreateFn createFn = nullptr;
    static_assert(sizeof(createFn) == sizeof(createProc));
    std::memcpy(&createFn, &createProc, sizeof(createFn));
    Check(createFn != nullptr, L"ViewerSqlite factory export is available for theme-cycle validation", success);
    if (! createFn)
    {
        return false;
    }

    wil::com_ptr<IViewer> viewer;
    const FactoryOptions factoryOptions{DEBUG_LEVEL_NONE};
    const HRESULT createHr = createFn(__uuidof(IViewer), &factoryOptions, nullptr, kViewerSqlitePluginId, viewer.put_void());
    Check(SUCCEEDED(createHr) && viewer != nullptr, L"ViewerSqlite factory creates an IViewer instance for theme-cycle validation", success);
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    const std::vector<HWND> existingWindows = CollectVisibleWindowsByClass(kViewerSqliteWindowClassName);

    BuiltinFileSystemStub fileSystem;
    const wchar_t* otherFiles[] = {databasePath.c_str()};
    ViewerOpenContext context{};
    context.fileSystem            = static_cast<IFileSystem*>(&fileSystem);
    context.fileSystemName        = L"File System";
    context.focusedPath           = databasePath.c_str();
    context.otherFiles            = otherFiles;
    context.otherFileCount        = static_cast<unsigned long>(std::size(otherFiles));
    context.focusedOtherFileIndex = 0u;

    const HRESULT openHr = viewer->Open(&context);
    Check(SUCCEEDED(openHr), L"viewer window open succeeds for theme-cycle validation", success);
    if (FAILED(openHr))
    {
        return false;
    }

    HWND viewerWindow       = nullptr;
    const bool openedWindow = PumpUntil(
        [&]() noexcept
    {
        viewerWindow = FindNewVisibleWindowByClass(kViewerSqliteWindowClassName, existingWindows);
        return viewerWindow != nullptr;
    },
        5000ms);
    Check(openedWindow, L"viewer window becomes visible for theme-cycle validation", success);
    if (! openedWindow || ! viewerWindow)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    WndMsg::ViewerSqliteDebugSnapshot snapshot{};
    const bool initialReady = WaitForViewerSnapshot(viewerWindow,
                                                    [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.tablePreviewMode && value.rowCount == 200u && value.rowOffset == 0u && value.firstRowPrimaryKey == 1u &&
               value.lastRowPrimaryKey == 200u && ! value.prevButtonEnabled && value.nextButtonEnabled && value.sortDirection == 0u;
    },
                                                    10000ms,
                                                    &snapshot);
    Check(initialReady, L"viewer window reaches the initial unsorted table-preview page for theme-cycle validation", success);
    if (! initialReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    const ViewerTheme initialTheme = MakeViewerTheme(
        TRUE, FALSE, FALSE, Argb(0x14, 0x14, 0x14), Argb(0xF2, 0xF2, 0xF2), Argb(0x23, 0x52, 0x7C), Argb(0xFF, 0xFF, 0xFF), Argb(0x57, 0xA4, 0xFF));
    Check(SUCCEEDED(viewer->SetTheme(&initialTheme)), L"viewer accepts the initial dark theme update", success);
    const bool initialThemeReady = WaitForViewerSnapshot(viewerWindow,
                                                         [&](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
    {
        return value.pendingAsyncWork == 0u && value.rowCount == 200u && value.rowOffset == 0u && value.themeDark == (initialTheme.darkMode != FALSE) &&
               value.themeHighContrast == (initialTheme.highContrast != FALSE) && value.themeRainbow == (initialTheme.rainbowMode != FALSE);
    },
                                                         5000ms,
                                                         &snapshot);
    Check(initialThemeReady, L"viewer window applies the initial dark theme for theme-cycle validation", success);
    if (! initialThemeReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    Check(DebugSelectViewerSqliteGridRow(viewerWindow, 1u), L"viewer debug row-selection hook can select the second row for theme-cycle validation", success);
    const bool selectionReady = WaitForViewerSnapshot(viewerWindow, [](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept {
        return value.selectionCount == 1u && value.primarySelectedRowId == 2u && value.selectedRowFillArgb != 0u && value.selectedRowTextArgb != 0u;
    }, 5000ms, &snapshot);
    Check(selectionReady, L"viewer window exposes a selected row for theme-cycle validation", success);
    if (! selectionReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    const auto requireVisibleUiaSelection = [&](std::wstring_view label, std::wstring* selectedNameOut) noexcept
    {
        UiaSelectionState selectionState{};
        const bool uiaReady = PumpUntil(
            [&]() noexcept
        {
            return TryCollectVisibleUiaViewerGridSelectionState(viewerWindow, selectionState) && selectionState.selectedCount == 1u &&
                   selectionState.selectedControlType == UIA_DataItemControlTypeId && selectionState.selectedHasSelectionItemPattern &&
                   selectionState.selectedVisible && ! selectionState.selectedName.empty();
        },
            5000ms);
        Check(uiaReady, std::format(L"viewer {} keeps a visible selected UIA DataItem with SelectionItemPattern", label), success);
        if (! uiaReady)
        {
            return false;
        }

        if (selectedNameOut)
        {
            *selectedNameOut = selectionState.selectedName;
        }
        return true;
    };

    std::wstring baselineSelectedName;
    const bool baselineUiaSelectionReady = requireVisibleUiaSelection(L"theme-cycle baseline selection", &baselineSelectedName);
    Check(baselineUiaSelectionReady, L"viewer theme-cycle baseline exposes a named selected UIA row", success);
    if (! baselineUiaSelectionReady)
    {
        static_cast<void>(viewer->Close());
        return false;
    }

    const auto unpackColor = [](uint32_t argb) noexcept
    {
        struct RgbaColor final
        {
            float r = 0.0f;
            float g = 0.0f;
            float b = 0.0f;
            float a = 0.0f;
        };

        return RgbaColor{static_cast<float>((argb >> 16u) & 0xFFu) / 255.0f,
                         static_cast<float>((argb >> 8u) & 0xFFu) / 255.0f,
                         static_cast<float>(argb & 0xFFu) / 255.0f,
                         static_cast<float>((argb >> 24u) & 0xFFu) / 255.0f};
    };
    const auto luminance = [&](uint32_t argb) noexcept
    {
        const auto color     = unpackColor(argb);
        const auto linearize = [](float channel) noexcept { return (channel <= 0.03928f) ? (channel / 12.92f) : std::pow((channel + 0.055f) / 1.055f, 2.4f); };

        return (0.2126f * linearize(color.r)) + (0.7152f * linearize(color.g)) + (0.0722f * linearize(color.b));
    };
    const auto contrastRatio = [&](uint32_t a, uint32_t b) noexcept
    {
        const float lumA    = luminance(a);
        const float lumB    = luminance(b);
        const float lighter = (std::max)(lumA, lumB);
        const float darker  = (std::min)(lumA, lumB);
        return (lighter + 0.05f) / (darker + 0.05f);
    };

    const auto requireTheme = [&](std::wstring_view label, const ViewerTheme& theme, const bool expectRainbow, const bool expectHighContrast) noexcept
    {
        const uint64_t previousRenderCount = snapshot.renderCount;
        Check(SUCCEEDED(viewer->SetTheme(&theme)), std::format(L"viewer accepts the {} theme update", label), success);
        if (! success)
        {
            return;
        }

        const bool themeReady = WaitForViewerSnapshot(viewerWindow,
                                                      [&](const WndMsg::ViewerSqliteDebugSnapshot& value) noexcept
        {
            return value.pendingAsyncWork == 0u && value.selectionCount == 1u && value.primarySelectedRowId == 2u &&
                   value.themeDark == (theme.darkMode != FALSE) && value.themeHighContrast == (theme.highContrast != FALSE) &&
                   value.themeRainbow == (theme.rainbowMode != FALSE) && value.selectedRowFillArgb != 0u && value.selectedRowTextArgb != 0u &&
                   value.renderCount >= previousRenderCount;
        },
                                                      5000ms,
                                                      &snapshot);
        Check(themeReady, std::format(L"viewer window settles after switching to the {} theme", label), success);
        if (! themeReady)
        {
            return;
        }

        Check(IsWindow(viewerWindow) != FALSE, std::format(L"viewer window survives the {} theme update", label), success);
        Check(snapshot.selectedRowUsesRainbow == expectRainbow, std::format(L"viewer selected-row rainbow state matches the {} theme", label), success);
        Check(snapshot.themeHighContrast == expectHighContrast, std::format(L"viewer high-contrast state matches the {} theme", label), success);
        Check(snapshot.selectedRowFillArgb != snapshot.selectedRowTextArgb,
              std::format(L"viewer selected-row colors stay distinct in the {} theme", label),
              success);

        const float minimumContrast = expectHighContrast ? 4.5f : 3.0f;
        Check(contrastRatio(snapshot.selectedRowFillArgb, snapshot.selectedRowTextArgb) >= minimumContrast,
              std::format(L"viewer selected-row text contrast stays above {:.1f}:1 in the {} theme", minimumContrast, label),
              success);

        std::wstring selectedName;
        const bool uiaSelectionReady = requireVisibleUiaSelection(std::format(L"{} theme selection", label), &selectedName);
        Check(uiaSelectionReady, std::format(L"viewer {} theme keeps a visible named selected UIA row", label), success);
        if (uiaSelectionReady)
        {
            Check(selectedName == baselineSelectedName, std::format(L"viewer {} theme preserves the selected UIA row name across theme churn", label), success);
        }
    };

    requireTheme(
        L"dark",
        MakeViewerTheme(
            TRUE, FALSE, FALSE, Argb(0x14, 0x14, 0x14), Argb(0xF2, 0xF2, 0xF2), Argb(0x23, 0x52, 0x7C), Argb(0xFF, 0xFF, 0xFF), Argb(0x57, 0xA4, 0xFF)),
        false,
        false);
    requireTheme(
        L"light",
        MakeViewerTheme(
            FALSE, FALSE, FALSE, Argb(0xFA, 0xFA, 0xFA), Argb(0x18, 0x18, 0x18), Argb(0xC9, 0xE1, 0xFF), Argb(0x10, 0x10, 0x10), Argb(0x1A, 0x73, 0xE8)),
        false,
        false);
    requireTheme(L"rainbow",
                 MakeViewerTheme(
                     TRUE, FALSE, TRUE, Argb(0x16, 0x11, 0x1C), Argb(0xF7, 0xF3, 0xFF), Argb(0x44, 0x2A, 0x8F), Argb(0xFF, 0xFF, 0xFF), Argb(0xFF, 0x7A, 0x18)),
                 true,
                 false);
    requireTheme(
        L"high-contrast",
        MakeViewerTheme(
            FALSE, TRUE, FALSE, Argb(0x00, 0x00, 0x00), Argb(0xFF, 0xFF, 0xFF), Argb(0xFF, 0xFF, 0x00), Argb(0x00, 0x00, 0x00), Argb(0x00, 0xFF, 0xFF)),
        false,
        true);

    const HRESULT closeHr = viewer->Close();
    Check(SUCCEEDED(closeHr), L"viewer window close succeeds after theme-cycle validation", success);
    Check(PumpUntil([&]() noexcept { return IsWindow(viewerWindow) == FALSE; }, 5000ms), L"viewer window closes cleanly after theme-cycle validation", success);

    static_cast<void>(viewer->SetCallback(nullptr, nullptr));
    viewer.reset();
    pluginModule.reset();
    return success;
}
#endif
} // namespace

int wmain(int argc, wchar_t** argv)
{
    std::wstring errorText;
    TempDatabase tempDb = CreateDatabase(errorText);
    if (tempDb.path.empty())
    {
        std::wcerr << std::format(L"Failed to create test database: {}\n", errorText);
        return 1;
    }

    const auto opened = ViewerSqliteEngine::DatabaseSource::OpenFromPath(tempDb.path, L"test.sqlite", false);
    if (FAILED(opened.hr) || ! opened.source)
    {
        std::wcerr << std::format(L"Failed to open test database: {}\n", opened.errorText);
        return 1;
    }

    const auto runSourceTest = [&](const wchar_t* name, auto&& test) noexcept
    {
        std::wcout << std::format(L"[ RUN      ] {}\n", name) << std::flush;
        const bool ok = test();
        std::wcout << std::format(L"[ {} ] {}\n", ok ? L"       OK" : L"  FAILED  ", name) << std::flush;
        return ok;
    };

    const auto runNamedSourceTest = [&](const wchar_t* name, const auto& test) noexcept
    { return runSourceTest(name, [&]() noexcept { return test(*opened.source); }); };

    const auto runNamedViewerTest = [&](const wchar_t* name, const auto& test) noexcept
    { return runSourceTest(name, [&]() noexcept { return test(tempDb.path); }); };

    const std::wstring filter = (argc >= 2 && argv != nullptr && argv[1] != nullptr) ? std::wstring(argv[1]) : std::wstring{};

    bool success         = true;
    const auto shouldRun = [&](std::wstring_view testName) noexcept { return filter.empty() || filter == testName; };

    if (shouldRun(L"TestListTables"))
    {
        success = runNamedSourceTest(L"TestListTables", TestListTables) && success;
    }
    if (shouldRun(L"TestPagedReads"))
    {
        success = runNamedSourceTest(L"TestPagedReads", TestPagedReads) && success;
    }
    if (shouldRun(L"TestSortedPagedReads"))
    {
        success = runNamedSourceTest(L"TestSortedPagedReads", TestSortedPagedReads) && success;
    }
    if (shouldRun(L"TestReadOnlyQueries"))
    {
        success = runNamedSourceTest(L"TestReadOnlyQueries", TestReadOnlyQueries) && success;
    }
    if (shouldRun(L"TestViewerWindowUsesDxUiHostWithNoVisibleChildControls"))
    {
        success =
            runNamedViewerTest(L"TestViewerWindowUsesDxUiHostWithNoVisibleChildControls", TestViewerWindowUsesDxUiHostWithNoVisibleChildControls) && success;
    }
#ifdef _DEBUG
    if (shouldRun(L"TestViewerWindowLongRunScrollingStaysBounded"))
    {
        success = runNamedViewerTest(L"TestViewerWindowLongRunScrollingStaysBounded", TestViewerWindowLongRunScrollingStaysBounded) && success;
    }
    if (shouldRun(L"TestViewerWindowLongRunOpenCloseStaysStable"))
    {
        success = runNamedViewerTest(L"TestViewerWindowLongRunOpenCloseStaysStable", TestViewerWindowLongRunOpenCloseStaysStable) && success;
    }
    if (shouldRun(L"TestViewerWindowPagingAndSortFlows"))
    {
        success = runNamedViewerTest(L"TestViewerWindowPagingAndSortFlows", TestViewerWindowPagingAndSortFlows) && success;
    }
#if ! defined(__SANITIZE_ADDRESS__)
    if (shouldRun(L"TestViewerWindowSelectionPatternSurvivesPagingAndSortFlows"))
    {
        success =
            runNamedViewerTest(L"TestViewerWindowSelectionPatternSurvivesPagingAndSortFlows", TestViewerWindowSelectionPatternSurvivesPagingAndSortFlows) &&
            success;
    }
#endif
    if (shouldRun(L"TestViewerWindowTabTraversalMatchesExpectedOrder"))
    {
        success = runNamedViewerTest(L"TestViewerWindowTabTraversalMatchesExpectedOrder", TestViewerWindowTabTraversalMatchesExpectedOrder) && success;
    }
    if (shouldRun(L"TestViewerWindowThemeCycleKeepsGridLegible"))
    {
        success = runNamedViewerTest(L"TestViewerWindowThemeCycleKeepsGridLegible", TestViewerWindowThemeCycleKeepsGridLegible) && success;
    }
#endif

    std::wcout << (success ? L"ViewerSqliteTests passed.\n" : L"ViewerSqliteTests failed.\n") << std::flush;

    if (! tempDb.path.empty())
    {
        std::error_code ec;
        std::filesystem::remove(tempDb.path, ec);
        tempDb.path.clear();
    }

    return success ? 0 : 1;
}

#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <objbase.h>

#include <atomic>
#include <chrono>
#include <condition_variable>
#include <cstddef>
#include <cstdint>
#include <deque>
#include <functional>
#include <memory>
#include <mutex>
#include <string>
#include <string_view>
#include <thread>
#include <unordered_map>
#include <utility>
#include <vector>

#include <d2d1.h>
#include <dwrite.h>

#pragma comment(lib, "d2d1")
#pragma comment(lib, "dwrite")

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/Terminal.h"
#include "TerminalCommandSurfaceModel.h"
#include "TerminalKittyImagePipeline.h"
#include "TerminalRuntimeLoader.h"
#if defined(ENABLE_TESTS)
#include "TerminalVtUpgradeTestContract.h"
#endif

[[nodiscard]] const char* GetTerminalStaticConfigurationSchema() noexcept;
[[nodiscard]] bool CanUnloadTerminalModuleNow() noexcept;

class TerminalAccessibility;

class Terminal final : public ITerminal, public ITerminalActions, public IInformations
{
public:
    Terminal();
    ~Terminal();

    Terminal(const Terminal&)            = delete;
    Terminal(Terminal&&)                 = delete;
    Terminal& operator=(const Terminal&) = delete;
    Terminal& operator=(Terminal&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** object) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    HRESULT STDMETHODCALLTYPE GetMetaData(const PluginMetaData** metaData) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfigurationSchema(const char** schemaJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SetConfiguration(const char* configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfiguration(const char** configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SomethingToSave(BOOL* somethingToSave) noexcept override;

    HRESULT STDMETHODCALLTYPE Open(const TerminalOpenContext* context) noexcept override;
    HRESULT STDMETHODCALLTYPE GetChildWindow(HWND* childWindow) noexcept override;
    HRESULT STDMETHODCALLTYPE SetTheme(const TerminalTheme* theme) noexcept override;
    HRESULT STDMETHODCALLTYPE UpdateSourceLocation(const TerminalSourceUpdate* update) noexcept override;
    HRESULT STDMETHODCALLTYPE InsertPath(const TerminalPathInsertion* insertion) noexcept override;
    HRESULT STDMETHODCALLTYPE GetViewState(TerminalViewState* state) noexcept override;
    HRESULT STDMETHODCALLTYPE SetCallback(ITerminalEventCallback* callback, void* cookie) noexcept override;
    HRESULT STDMETHODCALLTYPE Close() noexcept override;

    HRESULT STDMETHODCALLTYPE RouteShortcut(const TerminalShortcutRequest* request, TerminalShortcutRoute* route) noexcept override;
    HRESULT STDMETHODCALLTYPE GetActionState(const TerminalActionRequest* request, TerminalActionState* state) noexcept override;
    HRESULT STDMETHODCALLTYPE ExecuteAction(const TerminalActionRequest* request) noexcept override;

#if defined(ENABLE_TESTS)
    [[nodiscard]] HRESULT DebugGetDiagnosticText(TerminalOwnedUtf16* text) noexcept;
    [[nodiscard]] HRESULT DebugGetScreenText(TerminalOwnedUtf16* text) noexcept;
    [[nodiscard]] HRESULT DebugTerminateRootProcess(uint32_t exitCode) noexcept;
    [[nodiscard]] HRESULT DebugRunCommandExperiencePerfSelfTests(unsigned int* passedTests, unsigned int* failedTests) noexcept;
    [[nodiscard]] HRESULT DebugRunKittyCacheSelfTests(unsigned int* passedTests, unsigned int* failedTests) noexcept;
    [[nodiscard]] HRESULT DebugRunVtUpgradeCorpus(TerminalVtUpgradeTestContract::Evidence* evidence) noexcept;
    [[nodiscard]] HRESULT DebugRunInputRoutingContractSelfTests(unsigned int* passedTests, unsigned int* failedTests) noexcept;
    [[nodiscard]] HRESULT DebugRunRuntimeLifecycleSelfTests(unsigned int* passedTests, unsigned int* failedTests) noexcept;
    void DebugClearCapturedInput() noexcept;
    [[nodiscard]] std::string DebugTakeCapturedInput() noexcept;
    [[nodiscard]] DWORD DebugCommandSurfaceJoinTid() const noexcept;
    void DebugForceNextClipboardCopyFailure() noexcept;
    [[nodiscard]] size_t DebugCommandSurfaceRetiredWorkerCount() const noexcept;
#endif

private:
    using UniquePseudoConsole = wil::unique_any<HPCON, decltype(&::ClosePseudoConsole), ::ClosePseudoConsole>;

    enum class ShellKind : uint8_t
    {
        PowerShell,
        CommandPrompt,
        Posix,
    };

    enum class Osc52Policy : uint8_t
    {
        Deny,
        Ask,
        Allow,
    };

    enum class CommandSurfaceKind : uint8_t
    {
        None,
        Find,
        Suggestions,
    };

    struct TrustedPromptReply final
    {
        std::wstring followTarget;
        std::wstring promptReplacement;
        bool replacementAccepted = false;
        bool replacementRejected = false;
    };

    struct Config final
    {
        std::wstring fontFamily      = L"Cascadia Mono";
        float fontSizeDip            = 14.0f;
        uint32_t maxFormattedBytes   = 8u * 1024u * 1024u;
        uint32_t pasteMaxBytes       = 4u * 1024u * 1024u;
        uint32_t osc52MaxBytes       = 256u * 1024u;
        std::wstring defaultShell    = L"auto";
        std::wstring hyperlinkPolicy = L"ask";
        std::wstring osc52Policy     = L"ask";
        bool followPathWhenIdle      = false;
        bool warnOnUnsafePaste       = true;
    };

    struct Osc52ClipboardPayload final
    {
        ~Osc52ClipboardPayload() noexcept
        {
            SecureZeroMemory(utf8.data(), utf8.size());
        }

        std::string utf8;
    };

    struct ReadyWakePayload final
    {
        uint64_t sessionGeneration = 0u;
    };

    struct CommandSurfaceJoinContext final
    {
        CommandSurfaceJoinContext(std::jthread workerHandle, wil::unique_hmodule module, Terminal* value) noexcept
            : worker(std::move(workerHandle)),
              modulePin(std::move(module)),
              terminal(value)
        {
        }
        ~CommandSurfaceJoinContext()                                           = default;
        CommandSurfaceJoinContext(const CommandSurfaceJoinContext&)            = delete;
        CommandSurfaceJoinContext(CommandSurfaceJoinContext&&)                 = delete;
        CommandSurfaceJoinContext& operator=(const CommandSurfaceJoinContext&) = delete;
        CommandSurfaceJoinContext& operator=(CommandSurfaceJoinContext&&)      = delete;

        std::jthread worker;
        wil::unique_hmodule modulePin;
        Terminal* terminal;
    };

    struct InputDrainContext final
    {
        InputDrainContext(Terminal* value, wil::unique_hmodule module) noexcept : terminal(value), modulePin(std::move(module))
        {
        }
        ~InputDrainContext()                                   = default;
        InputDrainContext(const InputDrainContext&)            = delete;
        InputDrainContext(InputDrainContext&&)                 = delete;
        InputDrainContext& operator=(const InputDrainContext&) = delete;
        InputDrainContext& operator=(InputDrainContext&&)      = delete;

        Terminal* terminal;
        wil::unique_hmodule modulePin;
    };

    struct KittyImageKey final
    {
        uint32_t imageId         = 0u;
        uint64_t imageGeneration = 0u;

        bool operator==(const KittyImageKey&) const noexcept = default;
    };

    struct KittyImageKeyHash final
    {
        [[nodiscard]] size_t operator()(const KittyImageKey& key) const noexcept
        {
            const uint64_t mixed =
                key.imageGeneration ^ (static_cast<uint64_t>(key.imageId) + 0x9E3779B97F4A7C15ull + (key.imageGeneration << 6u) + (key.imageGeneration >> 2u));
            return std::hash<uint64_t>{}(mixed);
        }
    };

    enum class KittyImageState : uint8_t
    {
        Missing,
        EnginePending,
        Deferred,
        ConversionPending,
        Ready,
        Rejected,
    };

    enum class KittyBitmapUploadState : uint8_t
    {
        Eligible,
        RetryPending,
        RecreatePending,
        PermanentFailure,
    };

    enum class KittyBitmapUploadDisposition : uint8_t
    {
        Ready,
        Uploaded,
        DeferredBudget,
        DeferredBackoff,
        RetryScheduled,
        RecreateTarget,
        PermanentFailure,
        InvalidInput,
    };

    struct KittyBitmapUploadResult final
    {
        KittyBitmapUploadDisposition disposition = KittyBitmapUploadDisposition::InvalidInput;
        HRESULT hr                               = E_FAIL;
    };

    struct KittyImageSnapshot final
    {
        uint32_t width  = 0u;
        uint32_t height = 0u;
        std::vector<uint8_t> bgra;
        wil::com_ptr<ID2D1Bitmap> bitmap;
        KittyImageState state              = KittyImageState::Missing;
        uint64_t pendingRequestId          = 0u;
        uint64_t lastVisibleEpoch          = 0u;
        size_t deferredPinnedBytes         = 0u;
        size_t deferredVisibleKeyCount     = 0u;
        KittyBitmapUploadState uploadState = KittyBitmapUploadState::Eligible;
        HRESULT uploadHr                   = S_OK;
        ULONGLONG uploadRetryAtTick        = 0u;
        uint64_t uploadDeviceGeneration    = 0u;
        uint32_t uploadAttemptCount        = 0u;
    };

    struct KittyPlacementSnapshot final
    {
        KittyImageKey imageKey{};
        GhosttyKittyGraphicsPlacementRenderInfo renderInfo{};
        uint32_t xOffsetPixels = 0u;
        uint32_t yOffsetPixels = 0u;
        int32_t z              = 0;
    };

    struct MessageRouteResult final
    {
        bool handled   = false;
        LRESULT result = 0;
    };

    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    [[nodiscard]] MessageRouteResult routeInputMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    [[nodiscard]] MessageRouteResult routeMouseMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    [[nodiscard]] MessageRouteResult routeImeMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    [[nodiscard]] MessageRouteResult routeAccessibilityMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    [[nodiscard]] MessageRouteResult routeSessionMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    [[nodiscard]] MessageRouteResult routeClipboardMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    [[nodiscard]] MessageRouteResult routeRenderMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    [[nodiscard]] MessageRouteResult routeTeardownMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    static void GhosttyWritePty(GhosttyTerminal terminal, void* userData, const uint8_t* data, size_t length) noexcept;
    static void GhosttyClipboardWrite(GhosttyTerminal terminal, void* userData, const ::GhosttyClipboardWrite* write) noexcept;
    [[nodiscard]] GhosttyClipboardWriteResult handleGhosttyClipboardWrite(const ::GhosttyClipboardWrite* write) noexcept;
    static void CALLBACK CloseThreadpoolCallback(PTP_CALLBACK_INSTANCE instance, void* context, PTP_WORK work) noexcept;
    static void CALLBACK InputDrainCallback(PTP_CALLBACK_INSTANCE instance, void* context) noexcept;
    static void CALLBACK CommandSurfaceJoinCallback(PTP_CALLBACK_INSTANCE instance, void* context) noexcept;
    static void NotifyKittyReady(void* context) noexcept;

    [[nodiscard]] HRESULT createWindow(HWND parent) noexcept;
    void applyNativeScrollbarTheme(HWND hwnd) const noexcept;
    [[nodiscard]] HRESULT initializeEngine() noexcept;
    void releaseEngine() noexcept;
    [[nodiscard]] HRESULT startPseudoConsole(const TerminalLogicalLocation& location) noexcept;
    [[nodiscard]] HRESULT resolveShell(const TerminalLogicalLocation& location, std::wstring& executable, std::wstring& commandLine, ShellKind& kind) noexcept;
    void readerMain(std::stop_token stopToken, wil::unique_handle outputRead) noexcept;
    void processWatcherMain(std::stop_token stopToken, wil::unique_handle rootProcess) noexcept;
    void notifyOutputReady() noexcept;
    void postOutputReadyWake() noexcept;
    void consumeOutputReady() noexcept;
    void publishFinalExitEvent() noexcept;
    void integrationMain(std::stop_token stopToken, DWORD rootProcessId, std::wstring pipeName) noexcept;
    [[nodiscard]] TrustedPromptReply acceptTrustedPrompt(std::wstring path,
                                                         std::wstring historyPath,
                                                         std::wstring promptBuffer,
                                                         std::vector<std::wstring> currentSessionHistory,
                                                         bool supportsPromptReplacement) noexcept;
    void markUserInput(std::string_view bytes) noexcept;
    [[nodiscard]] bool writeInput(std::string_view bytes) noexcept;
    [[nodiscard]] bool writeTextInput(std::wstring_view text) noexcept;
    [[nodiscard]] bool pasteClipboard() noexcept;
    [[nodiscard]] bool encodePasteSource(std::string source) noexcept;
    [[nodiscard]] bool confirmationVisible() const noexcept;
    [[nodiscard]] std::wstring confirmationAccessibilityText() const;
    void resolveConfirmation(bool accept) noexcept;
    void drawConfirmation(float widthDip, float heightDip) noexcept;
    [[nodiscard]] bool handleConfirmationClick(POINT clientPoint) noexcept;
    [[nodiscard]] bool activateHyperlinkAt(POINT clientPoint) noexcept;
    [[nodiscard]] bool openHyperlink(std::wstring_view uri) noexcept;
    void handleOsc52Clipboard(std::unique_ptr<Osc52ClipboardPayload> payload) noexcept;
    [[nodiscard]] bool commitOsc52Clipboard(std::wstring_view text) noexcept;
    [[nodiscard]] bool openCommandSurface(CommandSurfaceKind kind) noexcept;
    void closeCommandSurface() noexcept;
    void restartCommandSurfaceQuery() noexcept;
    [[nodiscard]] bool handleCommandSurfaceKey(WPARAM virtualKey) noexcept;
    [[nodiscard]] bool handleCommandSurfaceCharacter(wchar_t character) noexcept;
    [[nodiscard]] bool handleCommandSurfaceClick(POINT clientPoint, bool activate) noexcept;
    void navigateCommandSurface(int direction) noexcept;
    void activateCommandSurfaceSelection() noexcept;
    void drawCommandSurface(float widthDip, float heightDip) noexcept;
    [[nodiscard]] std::wstring commandSurfaceAccessibilityText() const;
    void recordShortcutRoute(TerminalShortcutRoute route, HRESULT result, std::chrono::steady_clock::time_point startedAt) noexcept;
    void flushShortcutRouteMetrics() noexcept;
    [[nodiscard]] bool copySelection(bool* selectionPresent = nullptr) noexcept;
    [[nodiscard]] bool hasSelection() noexcept;
    void clearSelection() noexcept;
    [[nodiscard]] bool updateSelection(POINT clientPoint) noexcept;
    void stopSelection() noexcept;
    [[nodiscard]] bool selectWordAt(POINT clientPoint) noexcept;
    [[nodiscard]] bool selectAll() noexcept;
    [[nodiscard]] GhosttyPoint terminalPointFromClient(POINT clientPoint) const noexcept;
    [[nodiscard]] bool encodeMouse(GhosttyMouseAction action, GhosttyMouseButton button, POINT clientPoint, bool anyButtonPressed) noexcept;
    void releaseReportedMouseButtons() noexcept;
    void scrollViewport(intptr_t rows) noexcept;
    void scrollViewportTo(uint64_t row) noexcept;
    void setLiveFontSize(float fontSizeDip) noexcept;
    void handleVerticalScroll(UINT scrollCode) noexcept;
    [[nodiscard]] bool getScrollbarState(GhosttyTerminalScrollbar& state) noexcept;
    void updateScrollbar() noexcept;
    void updateImeCandidatePosition() noexcept;
    void drainInput() noexcept;
    void clearQueuedInput() noexcept;
    [[nodiscard]] bool encodeSpecialKey(WPARAM virtualKey, LPARAM keyData, bool released) noexcept;
    void resizeTerminal() noexcept;
    void render() noexcept;
    [[nodiscard]] HRESULT resizeRenderTargetToClient() noexcept;
    void handleResize(HWND hwnd) noexcept;
    [[nodiscard]] bool handleTimer(HWND hwnd, UINT_PTR timerId) noexcept;
    void requestKittyRepaint(HWND hwnd) noexcept;
    void armKittyUploadRetryTimer(ULONGLONG retryAtTick) noexcept;
    void scheduleKittyUploadRetry(KittyImageSnapshot& image, DWORD delayMilliseconds) noexcept;
    void drawInactivePaneOverlay(float widthDip, float heightDip, bool paneFocused) noexcept;
    [[nodiscard]] bool renderStructuredScreen(float widthDip, float heightDip) noexcept;
    [[nodiscard]] bool captureKittyGraphicsLocked(TerminalKittyGenerationWork& work,
                                                  std::vector<KittyPlacementSnapshot>& placements,
                                                  bool& submitWork) noexcept;
    [[nodiscard]] static bool handleKittyImageQueryResult(KittyImageSnapshot& image, GhosttyResult result) noexcept;
    static void resetKittyPendingForStorageChange(KittyImageSnapshot& image) noexcept;
    [[nodiscard]] static bool needsKittyReplacementWork(const KittyImageSnapshot& image, uint64_t currentRequestId, bool retryDeferred) noexcept;
    [[nodiscard]] static bool isCurrentKittyConversionResult(const KittyImageSnapshot& image, uint64_t requestId) noexcept;
    void consumeKittyReady() noexcept;
    void drawKittyLayer(GhosttyKittyPlacementLayer layer) noexcept;
    [[nodiscard]] KittyBitmapUploadResult ensureKittyBitmap(KittyImageSnapshot& image) noexcept;
    [[nodiscard]] bool recoverKittyDeviceAfterFrame(HWND hwnd) noexcept;
    void clearRenderDirtyState() noexcept;
    void publishAccessibilitySnapshot() noexcept;
    void discardDeviceResources() noexcept;
    [[nodiscard]] HRESULT ensureTextResources() noexcept;
    [[nodiscard]] HRESULT ensureDeviceResources() noexcept;
    void updateCellMetrics() noexcept;
    [[nodiscard]] std::wstring formatScreen() noexcept;
    void setDiagnostic(std::wstring text) noexcept;
    void prepareClose() noexcept;
    void finishClose() noexcept;
    [[nodiscard]] HRESULT ensureCloseWork() noexcept;
    void retireCommandSurfaceWorker() noexcept;
    void joinCommandSurfaceWorkers() noexcept;
    [[nodiscard]] uint32_t commandSurfaceJoinsOutstanding() noexcept;
    void setLifecycle(TerminalLifecycleState lifecycle) noexcept;
    [[nodiscard]] static bool validateLocation(const TerminalLogicalLocation& location) noexcept;
    [[nodiscard]] static std::wstring copySpan(TerminalUtf16Span span);
    [[nodiscard]] static HRESULT copyOwned(std::wstring_view text, TerminalOwnedUtf16& output) noexcept;
    [[nodiscard]] std::wstring quotePath(std::wstring_view path) const;

    std::atomic<ULONG> _refCount{1u};
    std::atomic<HWND> _windowHandle{nullptr};
    wil::unique_hwnd _window;

    TerminalRuntimeLoader _runtime;
    GhosttyTerminal _ghosttyTerminal                              = nullptr;
    GhosttyFormatter _formatter                                   = nullptr;
    GhosttyRenderState _renderState                               = nullptr;
    GhosttyRenderStateRowIterator _renderRows                     = nullptr;
    GhosttyRenderStateRowCells _renderCells                       = nullptr;
    GhosttyKeyEncoder _keyEncoder                                 = nullptr;
    GhosttyKeyEvent _keyEvent                                     = nullptr;
    GhosttyMouseEncoder _mouseEncoder                             = nullptr;
    GhosttyMouseEvent _mouseEvent                                 = nullptr;
    GhosttyKittyGraphicsPlacementIterator _kittyPlacementIterator = nullptr;
    std::mutex _terminalMutex;
    std::wstring _diagnosticText;
    std::wstring _engineInitializationError;
    std::unique_ptr<TerminalAccessibility> _accessibility;

    UniquePseudoConsole _pseudoConsole;
    wil::unique_handle _ptyInputWrite;
    wil::unique_handle _processJob;
    wil::unique_handle _process;
    std::jthread _readerThread;
    std::jthread _processWatcherThread;
    std::jthread _integrationThread;
    std::jthread _commandSurfaceWorker;
    std::vector<std::jthread> _retiredCommandSurfaceWorkers;
    std::mutex _commandSurfaceJoinMutex;
    std::condition_variable _commandSurfaceJoinCv;
    uint32_t _commandSurfaceJoinOutstanding = 0u;
#if defined(ENABLE_TESTS)
    std::atomic<DWORD> _debugCommandSurfaceJoinTid{0};
    std::atomic_bool _debugForceNextClipboardCopyFailure{false};
    std::atomic_uint32_t _debugEngineInitializationFailStage{0u};
    std::atomic_uint32_t _debugClipboardCallbackCount{0u};
    std::atomic_uint32_t _debugClipboardReplyCount{0u};
    std::atomic<GhosttyClipboardWriteResult> _debugClipboardReplyResult{GHOSTTY_CLIPBOARD_WRITE_RESULT_MAX_VALUE};
    std::atomic_uint32_t _debugOsc52PayloadPostCount{0u};
#endif
    wil::unique_threadpool_work_nowait _closeWork;
    wil::unique_hmodule _closeModulePin;
    std::mutex _ioMutex;
    std::mutex _sessionTeardownMutex;
    std::mutex _inputMutex;
    std::deque<std::string> _inputQueue;
    size_t _inputQueuedBytes    = 0u;
    size_t _inputActiveRequests = 0u;
    bool _inputDrainScheduled   = false;
    std::atomic_uint64_t _outputGeneration{0u};
    std::atomic_uint64_t _outputConsumedGeneration{0u};
    std::atomic_bool _outputNotificationPending{false};
    std::atomic_bool _closing{false};
    std::atomic_bool _exitCodePresent{false};
    std::atomic_uint32_t _exitCode{0u};
    std::atomic_bool _finalSnapshotComplete{false};
    std::atomic_bool _rootProcessExited{false};
    std::atomic_int64_t _rootExitObservedTimestampNs{0};
    std::atomic_int64_t _lastInputTimestampNs{0};
    std::atomic_int64_t _lastOutputTimestampNs{0};
    std::atomic_bool _followPathWhenIdle{false};
    ShellKind _shellKind = ShellKind::PowerShell;

    std::mutex _stateMutex;
    TerminalOriginalSourceKey _originalSource{};
    uint64_t _sourceGeneration               = 0u;
    TerminalLocationKind _sourceLocationKind = TerminalLocationKind::Unsupported;
    std::wstring _sourceLocationPath;
    std::wstring _sourcePluginShortId;
    std::wstring _launchWslDistribution;
    std::wstring _trustedCurrentDirectory;
    std::wstring _trustedHistoryPath;
    std::wstring _trustedPromptBuffer;
    std::vector<std::wstring> _trustedCurrentSessionHistory;
    std::wstring _pendingPromptExpectedBuffer;
    std::wstring _pendingPromptReplacement;
    std::wstring _pendingFollowDirectory;
    TerminalFollowState _followState      = TerminalFollowState::Disabled;
    bool _integrationTrusted              = false;
    bool _trustedHistoryProviderAvailable = false;
    bool _idleAtPrimaryPrompt             = false;
    bool _hasPendingUserInput             = false;
    bool _commandSubmitted                = false;
#if defined(ENABLE_TESTS)
    std::atomic_uint32_t _integrationDebugStage{0u};
    std::atomic<HRESULT> _integrationDebugResult{S_OK};
    std::atomic_uint32_t _integrationDebugClientPid{0u};
    uint64_t _debugScrollViewportMutationCount = 0u;
#endif

    wil::com_ptr<ID2D1Factory> _d2dFactory;
    wil::com_ptr<IDWriteFactory> _dwriteFactory;
    wil::com_ptr<ID2D1HwndRenderTarget> _renderTarget;
    wil::com_ptr<ID2D1SolidColorBrush> _foregroundBrush;
    wil::com_ptr<IDWriteTextFormat> _textFormat;
    wil::com_ptr<IDWriteTextFormat> _boldTextFormat;
    wil::com_ptr<IDWriteTextFormat> _italicTextFormat;
    wil::com_ptr<IDWriteTextFormat> _boldItalicTextFormat;
    wil::com_ptr<IDWriteTextFormat> _overlayTextFormat;
    wil::com_ptr<IDWriteTextFormat> _overlayBoldTextFormat;
    TerminalKittyImagePipeline _kittyPipeline;
    std::unordered_map<KittyImageKey, KittyImageSnapshot, KittyImageKeyHash> _kittyImages;
    std::vector<KittyPlacementSnapshot> _kittyPlacements;
    uint64_t _kittyRequestId                   = 0u;
    uint64_t _kittyTargetStorageGeneration     = 0u;
    uint64_t _kittyCacheEpoch                  = 0u;
    size_t _kittyCachedBytes                   = 0u;
    size_t _kittyCaptureSourceBytes            = 0u;
    size_t _kittyCapturePinnedBytes            = 0u;
    size_t _kittyCaptureEnginePlacementBytes   = 0u;
    size_t _kittyCaptureEnginePlacementCount   = 0u;
    size_t _kittyCaptureTotalPlacements        = 0u;
    size_t _kittyCaptureVisiblePlacements      = 0u;
    size_t _kittyCaptureOffscreenPlacements    = 0u;
    size_t _kittyCaptureVirtualPlacements      = 0u;
    size_t _kittyCaptureVisibleImageKeys       = 0u;
    size_t _kittyCaptureMissingImageKeys       = 0u;
    size_t _kittyCaptureEnginePendingImageKeys = 0u;
    size_t _kittyCaptureDeferredImageKeys      = 0u;
    size_t _kittyCaptureRejectedImageKeys      = 0u;
    size_t _kittySupersededImageCount          = 0u;
    size_t _kittyRemovedImageCount             = 0u;
    size_t _kittyEvictedBytes                  = 0u;
    size_t _kittyEvictedImageCount             = 0u;
    size_t _kittyFrameUploadBytes              = 0u;
    uint32_t _kittyFrameUploadCount            = 0u;
    size_t _kittyFrameImageLookupCount         = 0u;
    bool _kittyUploadDeferred                  = false;
    bool _kittyRecreateTargetAfterFrame        = false;
    ULONGLONG _kittyNextUploadRetryAtTick      = 0u;
    uint64_t _kittyDeviceGeneration            = 0u;
    uint32_t _kittyUploadRetryScheduleCount    = 0u;
    uint32_t _kittyStableUploadFailureCount    = 0u;
    uint32_t _kittyDeviceRecoveryCount         = 0u;
    uint32_t _kittyRepaintRequestCount         = 0u;
    uint32_t _kittyResizeCount                 = 0u;
    uint32_t _kittyResizeRecreateCount         = 0u;
    size_t _kittyResizeRetainedBitmapCount     = 0u;
    size_t _kittyResizeRetainedBitmapBytes     = 0u;
#if defined(ENABLE_TESTS)
    std::deque<HRESULT> _debugKittyUploadFailures;
#endif
    uint32_t _backgroundArgb               = 0xFF101010u;
    uint32_t _foregroundArgb               = 0xFFF0F0F0u;
    uint32_t _cursorArgb                   = 0xFFF0F0F0u;
    uint32_t _selectionBackgroundArgb      = 0xFF264F78u;
    uint32_t _selectionForegroundArgb      = 0xFFFFFFFFu;
    uint32_t _hyperlinkArgb                = 0xFF4EA1FFu;
    bool _themeDark                        = true;
    bool _themeHighContrast                = false;
    UINT _dpi                              = USER_DEFAULT_SCREEN_DPI;
    float _cellWidthDip                    = 8.68f;
    float _cellHeightDip                   = 19.6f;
    float _liveFontSizeDip                 = 14.0f;
    bool _cellMetricsValid                 = false;
    uint16_t _columns                      = 80u;
    uint16_t _rows                         = 24u;
    bool _ptyClientSizeSynced              = false;
    int _mouseWheelDeltaRemainder          = 0;
    GhosttyTrackedGridRef _selectionAnchor = nullptr;
    POINT _selectionLastPoint{};
    int _selectionAutoScrollRows  = 0;
    bool _selecting               = false;
    uint8_t _reportedMouseButtons = 0u;
    POINT _reportedMouseLastPoint{};
    wchar_t _pendingHighSurrogate = L'\0';
    std::wstring _imeComposition;
    std::string _pendingUnsafePaste;
    std::wstring _pendingHyperlink;
    std::wstring _pendingOsc52Text;
    bool _unsafePasteConfirmationVisible = false;
    bool _hyperlinkConfirmationVisible   = false;
    bool _osc52ConfirmationVisible       = false;

    CommandSurfaceKind _commandSurfaceKind = CommandSurfaceKind::None;
    uint64_t _commandSurfaceGeneration     = 0u;
    bool _commandSurfaceLoading            = false;
    std::wstring _commandSurfaceQuery;
    std::shared_ptr<std::wstring> _commandSurfaceSnapshot;
    std::shared_ptr<std::vector<std::wstring>> _commandSurfaceSessionHistorySnapshot;
    std::wstring _commandSurfacePromptSnapshot;
    std::wstring _commandSurfaceHistoryPath;
    std::wstring _persistedHistoryPath;
    std::wstring _commandSurfaceStatusOverride;
    uint64_t _commandSurfaceSessionGeneration  = 0u;
    uint64_t _commandSurfaceStateGeneration    = 0u;
    bool _commandSurfaceHistoryTrusted         = false;
    bool _commandSurfaceInsertionPending       = false;
    uint64_t _commandSurfaceSnapshotGeneration = 0u;
    bool _commandSurfaceSnapshotReady          = false;
    std::vector<TerminalCommandSurfaceModel::FindMatch> _commandSurfaceFindRows;
    size_t _commandSurfaceFindTotal = 0u;
    std::vector<TerminalCommandSurfaceModel::Suggestion> _commandSurfaceSuggestionRows;
    std::shared_ptr<std::vector<std::wstring>> _persistedHistory = std::make_shared<std::vector<std::wstring>>();
    size_t _commandSurfaceSelected                               = 0u;
    uint64_t _shortcutRouteBatchCount                            = 0u;
    uint64_t _shortcutRouteBatchTotalUs                          = 0u;
    uint64_t _shortcutRouteBatchMaximumUs                        = 0u;
    uint64_t _shortcutRouteBatchHandled                          = 0u;
    uint64_t _shortcutRouteBatchPassThrough                      = 0u;
    uint64_t _shortcutRouteBatchHost                             = 0u;
    uint64_t _shortcutRouteBatchBlocked                          = 0u;
    uint64_t _shortcutRouteBatchFailed                           = 0u;
    std::atomic<Osc52Policy> _osc52Policy{Osc52Policy::Ask};
    std::atomic_uint32_t _osc52MaxBytes{256u * 1024u};
    std::atomic_bool _osc52PostPending{false};
    bool _translatedCharacterSuppressionPending = false;
#if defined(ENABLE_TESTS)
    std::string _debugCapturedInput;
#endif
    std::wstring _unsafePasteTitle;
    std::wstring _unsafePasteMessage;
    std::wstring _unsafePasteAccept;
    std::wstring _unsafePasteCancel;
    std::wstring _hyperlinkTitle;
    std::wstring _hyperlinkMessage;
    std::wstring _hyperlinkAccept;
    std::wstring _hyperlinkCancel;
    std::wstring _osc52Title;
    std::wstring _osc52Message;
    std::wstring _osc52Accept;
    std::wstring _osc52Cancel;

    TerminalInstanceId _instanceId{};
    std::atomic<TerminalLifecycleState> _lifecycle{TerminalLifecycleState::Created};
    std::atomic_uint64_t _stateGeneration{1u};
    std::atomic_uint64_t _sessionGeneration{0u};

    std::mutex _eventCallbackMutex;
    ITerminalEventCallback* _eventCallback    = nullptr; // Weak; host owns it.
    void* _eventCallbackCookie                = nullptr;
    uint64_t _eventDeliveredSessionGeneration = 0u;

    Config _config;
    std::string _configurationJson;
    std::wstring _metaId;
    std::wstring _metaShortId;
    std::wstring _metaName;
    std::wstring _metaDescription;
    PluginMetaData _metaData{};
};

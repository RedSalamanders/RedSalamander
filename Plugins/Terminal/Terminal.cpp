#include "Terminal.h"
#include "TerminalAccessibility.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <clocale>
#include <cmath>
#include <cstdlib>
#include <cstring>
#include <format>
#include <limits>
#include <new>
#include <numeric>
#include <optional>
#include <ranges>
#include <set>
#include <system_error>
#include <utility>
#include <vector>

#include <windowsx.h>
#include <wincrypt.h>
#include <imm.h>
#include <shlobj.h>
#include <shellapi.h>
#include <UIAutomation.h>
#include <uxtheme.h>

#pragma comment(lib, "crypt32")
#pragma comment(lib, "imm32")
#pragma comment(lib, "advapi32")
#pragma comment(lib, "shell32")
#pragma comment(lib, "uxtheme")

#include <yyjson.h>

#include "Helpers.h"
#include "DxUi/DxUi.h"
#include "PathUtils.h"
#include "PaneVisualState.h"
#include "ProcessCommandLine.h"
#include "StringConversion.h"
#include "UnicodeClipboard.h"
#include "WindowMessages.h"
#include "resource.h"

extern HINSTANCE g_hInstance;

#include "TerminalInternal.h"

using namespace TerminalPluginDetail;


const char* GetTerminalStaticConfigurationSchema() noexcept
{
    std::string schema = BuildTerminalConfigurationSchema();
    if (schema.empty())
    {
        return nullptr;
    }
    std::scoped_lock lock(g_terminalSchemaMutex);
    const auto stored = g_terminalSchemas.insert(std::move(schema)).first;
    return stored->c_str();
}

bool CanUnloadTerminalModuleNow() noexcept
{
    if (g_terminalObjectCount.load(std::memory_order_acquire) != 0u ||
        g_terminalWindowCount.load(std::memory_order_acquire) != 0u || ! CanUnloadTerminalAccessibilityProviders())
    {
        return false;
    }

    std::scoped_lock lock(g_terminalWindowClassMutex);
    if (g_terminalWindowClassRegistered)
    {
        if (UnregisterClassW(kTerminalWindowClass, g_hInstance) == FALSE)
        {
            return false;
        }
        g_terminalWindowClassRegistered = false;
    }
    return true;
}

#ifdef ENABLE_TESTS
HRESULT Terminal::DebugGetDiagnosticText(TerminalOwnedUtf16* text) noexcept
{
    if (text == nullptr)
    {
        return E_POINTER;
    }
    std::scoped_lock lock(_terminalMutex);
    return copyOwned(_diagnosticText, *text);
}

HRESULT Terminal::DebugTerminateRootProcess(uint32_t exitCode) noexcept
{
    if (!_process)
    {
        return E_HANDLE;
    }
    return TerminateProcess(_process.get(), exitCode) != FALSE ? S_OK : HRESULT_FROM_WIN32(GetLastError());
}

HRESULT Terminal::DebugGetScreenText(TerminalOwnedUtf16* text) noexcept
{
    if (text == nullptr)
    {
        return E_POINTER;
    }
    // Reuse the bounded formatter used by the accessibility document. Host
    // lifecycle tests need real VT output without UIA broker discovery.
    return copyOwned(formatScreen(), *text);
}

#if defined(ENABLE_TESTS)
void Terminal::DebugClearCapturedInput() noexcept
{
    std::unique_lock lock(_inputMutex);
    _debugCapturedInput.clear();
}

std::string Terminal::DebugTakeCapturedInput() noexcept
{
    std::unique_lock lock(_inputMutex);
    std::string captured;
    captured.swap(_debugCapturedInput);
    return captured;
}

DWORD Terminal::DebugCommandSurfaceJoinTid() const noexcept
{
    return _debugCommandSurfaceJoinTid.load(std::memory_order_acquire);
}

void Terminal::DebugForceNextClipboardCopyFailure() noexcept
{
    _debugForceNextClipboardCopyFailure.store(true, std::memory_order_release);
}

size_t Terminal::DebugCommandSurfaceRetiredWorkerCount() const noexcept
{
    return _retiredCommandSurfaceWorkers.size();
}
#endif

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderTerminalDebugGetDiagnosticText(
    ITerminal* terminal,
    TerminalOwnedUtf16* text) noexcept
{
    if (terminal == nullptr || text == nullptr)
    {
        return E_POINTER;
    }
    return static_cast<Terminal*>(terminal)->DebugGetDiagnosticText(text);
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderTerminalDebugTerminateRootProcess(
    ITerminal* terminal,
    uint32_t exitCode) noexcept
{
    if (terminal == nullptr)
    {
        return E_POINTER;
    }
    return static_cast<Terminal*>(terminal)->DebugTerminateRootProcess(exitCode);
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderTerminalDebugGetScreenText(
    ITerminal* terminal,
    TerminalOwnedUtf16* text) noexcept
{
    if (terminal == nullptr || text == nullptr)
    {
        return E_POINTER;
    }
    return static_cast<Terminal*>(terminal)->DebugGetScreenText(text);
}

#if defined(ENABLE_TESTS)
extern "C" __declspec(dllexport) DWORD __stdcall RedSalamanderTerminalDebugCommandSurfaceJoinTid(
    ITerminal* terminal) noexcept
{
    if (terminal == nullptr)
    {
        return 0u;
    }
    return static_cast<Terminal*>(terminal)->DebugCommandSurfaceJoinTid();
}
#endif

#if defined(ENABLE_TESTS)
HRESULT Terminal::DebugRunInputRoutingContractSelfTests(
    unsigned int* passedTests,
    unsigned int* failedTests) noexcept
{
    if (passedTests == nullptr || failedTests == nullptr)
    {
        return E_POINTER;
    }
    *passedTests = 0u;
    *failedTests = 0u;
    const auto check = [passedTests, failedTests](bool condition, std::wstring_view name) noexcept
    {
        if (condition)
        {
            ++*passedTests;
            return;
        }
        ++*failedTests;
        Debug::Error(L"Terminal input-routing contract failed: {}", name);
    };

    const HRESULT initializeHr = initializeEngine();
    check(SUCCEEDED(initializeHr), L"initializeEngine");
    if (FAILED(initializeHr))
    {
        return E_FAIL;
    }

    DebugClearCapturedInput();
    static_cast<void>(ConsumeTranslatedCharacterSuppression(_translatedCharacterSuppressionPending));
    const auto enterKeydown = routeInputMessage(nullptr, WM_KEYDOWN, VK_RETURN, 1);
    const auto enterChar = routeInputMessage(nullptr, WM_CHAR, static_cast<WPARAM>(L'\r'), 1);
    const std::string enterBytes = DebugTakeCapturedInput();
    const size_t enterCarriageReturns = static_cast<size_t>(std::ranges::count(enterBytes, '\r'));
    check(enterKeydown.handled && enterChar.handled && enterCarriageReturns == 1u &&
              enterBytes.find("\r\r") == std::string::npos,
          std::format(L"Enter single CR (handled={}, charHandled={}, crs={}, bytes={})",
                      enterKeydown.handled,
                      enterChar.handled,
                      enterCarriageReturns,
                      enterBytes.size()));

    DebugClearCapturedInput();
    static_cast<void>(ConsumeTranslatedCharacterSuppression(_translatedCharacterSuppressionPending));
    const auto upKeydown = routeInputMessage(nullptr, WM_KEYDOWN, VK_UP, 1);
    const auto letterA = routeInputMessage(nullptr, WM_CHAR, static_cast<WPARAM>(L'a'), 1);
    const std::string upThenA = DebugTakeCapturedInput();
    check(upKeydown.handled && letterA.handled && upThenA.find('a') != std::string::npos,
          std::format(L"Up Arrow then a (handled={}, charHandled={}, bytes={}, hasA={})",
                      upKeydown.handled,
                      letterA.handled,
                      upThenA.size(),
                      upThenA.find('a') != std::string::npos));

    DebugClearCapturedInput();
    static_cast<void>(ConsumeTranslatedCharacterSuppression(_translatedCharacterSuppressionPending));
    _unsafePasteConfirmationVisible = true;
    const auto confirmationEscapeKeydown = routeInputMessage(nullptr, WM_KEYDOWN, VK_ESCAPE, 1);
    const auto confirmationEscapeChar = routeInputMessage(nullptr, WM_CHAR, static_cast<WPARAM>(L'\x1B'), 1);
    const std::string confirmationEscapeBytes = DebugTakeCapturedInput();
    check(confirmationEscapeKeydown.handled && confirmationEscapeChar.handled && ! confirmationVisible() &&
              confirmationEscapeBytes.empty(),
          std::format(L"confirmation Escape suppression (keydownHandled={}, charHandled={}, visible={}, bytes={})",
                      confirmationEscapeKeydown.handled,
                      confirmationEscapeChar.handled,
                      confirmationVisible(),
                      confirmationEscapeBytes.size()));

    DebugClearCapturedInput();
    static_cast<void>(ConsumeTranslatedCharacterSuppression(_translatedCharacterSuppressionPending));
    check(openCommandSurface(CommandSurfaceKind::Suggestions), L"openCommandSurface(Suggestions) for Escape suppression");
    const auto surfaceEscapeKeydown = routeInputMessage(nullptr, WM_KEYDOWN, VK_ESCAPE, 1);
    const auto surfaceEscapeChar = routeInputMessage(nullptr, WM_CHAR, static_cast<WPARAM>(L'\x1B'), 1);
    const std::string surfaceEscapeBytes = DebugTakeCapturedInput();
    check(surfaceEscapeKeydown.handled && surfaceEscapeChar.handled &&
              _commandSurfaceKind == CommandSurfaceKind::None && surfaceEscapeBytes.empty(),
          std::format(L"command-surface Escape suppression (keydownHandled={}, charHandled={}, kind={}, bytes={})",
                      surfaceEscapeKeydown.handled,
                      surfaceEscapeChar.handled,
                      static_cast<unsigned>(_commandSurfaceKind),
                      surfaceEscapeBytes.size()));

    constexpr std::string_view kSelectionFixture = "copy-selection-fixture\r\n";
    if (_runtime.terminalVtWrite != nullptr && _ghosttyTerminal != nullptr)
    {
        _runtime.terminalVtWrite(
            _ghosttyTerminal,
            reinterpret_cast<const uint8_t*>(kSelectionFixture.data()),
            kSelectionFixture.size());
    }
    const bool selectedAll = selectAll();
    check(selectedAll, L"selectAll");
    check(hasSelection(), L"hasSelection after selectAll");
    DebugForceNextClipboardCopyFailure();
    TerminalShortcutRequest copyRequest{};
    copyRequest.sizeBytes = sizeof(copyRequest);
    copyRequest.message = WM_KEYDOWN;
    copyRequest.virtualKey = VK_RETURN;
    copyRequest.instanceId = _instanceId;
    copyRequest.sessionGeneration = _sessionGeneration.load(std::memory_order_acquire);
    constexpr std::wstring_view kCopyOrPassCommand = L"cmd/terminal/copySelectionOrPassthrough";
    copyRequest.commandId = {kCopyOrPassCommand.data(), static_cast<uint32_t>(kCopyOrPassCommand.size())};
    TerminalShortcutRoute copyRoute = TerminalShortcutRoute::PassThrough;
    const HRESULT copyHr = RouteShortcut(&copyRequest, &copyRoute);
    check(copyHr == S_OK && copyRoute == TerminalShortcutRoute::Handled && hasSelection(),
          std::format(L"failed clipboard copy remains Handled (hr=0x{:08X}, route={}, hasSelection={})",
                      static_cast<unsigned long>(copyHr),
                      static_cast<unsigned>(copyRoute),
                      hasSelection()));

    check(openCommandSurface(CommandSurfaceKind::Suggestions), L"openCommandSurface(Suggestions)");
    for (int queryIndex = 0; queryIndex < 12; ++queryIndex)
    {
        _commandSurfaceQuery = L"x";
        restartCommandSurfaceQuery();
    }
    const auto joinDeadline = std::chrono::steady_clock::now() + std::chrono::seconds{2};
    while (commandSurfaceJoinsOutstanding() != 0u &&
           std::chrono::steady_clock::now() < joinDeadline)
    {
        Sleep(1u);
    }
    const size_t retiredWorkers = DebugCommandSurfaceRetiredWorkerCount();
    const uint32_t outstandingJoins = commandSurfaceJoinsOutstanding();
    check(outstandingJoins == 0u && retiredWorkers <= 2u,
          std::format(L"command-surface reaper (outstanding={}, retired={})", outstandingJoins, retiredWorkers));
    closeCommandSurface();
    const auto closeJoinDeadline = std::chrono::steady_clock::now() + std::chrono::seconds{2};
    while (commandSurfaceJoinsOutstanding() != 0u &&
           std::chrono::steady_clock::now() < closeJoinDeadline)
    {
        Sleep(1u);
    }
    check(commandSurfaceJoinsOutstanding() == 0u, L"closeCommandSurface retired worker reached the quiet point");

    constexpr std::string_view clipboardMime = "text/plain;charset=utf-8";
    constexpr std::string_view clipboardText = "clipboard-policy";
    GhosttyClipboardContent clipboardContent{};
    clipboardContent.mime.ptr = reinterpret_cast<const uint8_t*>(clipboardMime.data());
    clipboardContent.mime.len = clipboardMime.size();
    clipboardContent.data.ptr = reinterpret_cast<const uint8_t*>(clipboardText.data());
    clipboardContent.data.len = clipboardText.size();
    ::GhosttyClipboardWrite clipboardWrite{};
    clipboardWrite.size = sizeof(clipboardWrite);
    clipboardWrite.location = GHOSTTY_CLIPBOARD_LOCATION_STANDARD;
    clipboardWrite.contents = &clipboardContent;
    clipboardWrite.contents_len = 1u;

    _closing.store(false, std::memory_order_release);
    _osc52PostPending.store(false, std::memory_order_release);
    _osc52Policy.store(Osc52Policy::Deny, std::memory_order_release);
    check(handleGhosttyClipboardWrite(&clipboardWrite) == GHOSTTY_CLIPBOARD_WRITE_RESULT_DENIED,
          L"OSC 52 deny policy replies denied");

    _osc52PostPending.store(true, std::memory_order_release);
    _osc52Policy.store(Osc52Policy::Ask, std::memory_order_release);
    check(handleGhosttyClipboardWrite(&clipboardWrite) == GHOSTTY_CLIPBOARD_WRITE_RESULT_BUSY,
          L"OSC 52 ask policy rejects overlapping UI work as busy");
    _osc52Policy.store(Osc52Policy::Allow, std::memory_order_release);
    check(handleGhosttyClipboardWrite(&clipboardWrite) == GHOSTTY_CLIPBOARD_WRITE_RESULT_BUSY,
          L"OSC 52 allow policy rejects overlapping UI work as busy");

    constexpr std::array<uint8_t, 2u> invalidUtf8{{0xC3u, 0x28u}};
    _osc52PostPending.store(false, std::memory_order_release);
    clipboardContent.data.ptr = invalidUtf8.data();
    clipboardContent.data.len = invalidUtf8.size();
    check(handleGhosttyClipboardWrite(&clipboardWrite) == GHOSTTY_CLIPBOARD_WRITE_RESULT_INVALID_DATA,
          L"OSC 52 callback rejects invalid UTF-8 before posting UI work");

    clipboardContent.data.ptr = reinterpret_cast<const uint8_t*>(clipboardText.data());
    clipboardContent.data.len = clipboardText.size();
    _closing.store(true, std::memory_order_release);
    check(handleGhosttyClipboardWrite(&clipboardWrite) == GHOSTTY_CLIPBOARD_WRITE_RESULT_DENIED,
          L"OSC 52 callback denies teardown-before-UI work");
    _closing.store(false, std::memory_order_release);
    _osc52PostPending.store(false, std::memory_order_release);
    _osc52Policy.store(Osc52Policy::Ask, std::memory_order_release);

    return *failedTests == 0u ? S_OK : E_FAIL;
}

HRESULT Terminal::DebugRunRuntimeLifecycleSelfTests(
    unsigned int* passedTests,
    unsigned int* failedTests) noexcept
{
    if (passedTests == nullptr || failedTests == nullptr)
    {
        return E_POINTER;
    }
    *passedTests = 0u;
    *failedTests = 0u;
    const auto check = [passedTests, failedTests](bool condition, std::wstring_view name) noexcept
    {
        if (condition)
        {
            ++*passedTests;
            return;
        }
        ++*failedTests;
        Debug::Error(L"Terminal runtime lifecycle contract failed: {}", name);
    };
    const auto engineReleased = [](const Terminal& terminal) noexcept
    {
        return ! terminal._runtime.IsLoaded() && terminal._ghosttyTerminal == nullptr && terminal._formatter == nullptr &&
            terminal._renderState == nullptr && terminal._renderRows == nullptr && terminal._renderCells == nullptr &&
            terminal._keyEncoder == nullptr && terminal._keyEvent == nullptr && terminal._mouseEncoder == nullptr &&
            terminal._mouseEvent == nullptr && terminal._kittyPlacementIterator == nullptr &&
            terminal._kittyPipeline.GetStats().activeWorkers == 0u;
    };

    constexpr uint32_t kInitializationStageCount = 24u;
    for (uint32_t stage = 1u; stage <= kInitializationStageCount; ++stage)
    {
        Terminal terminal;
        terminal._debugEngineInitializationFailStage.store(stage, std::memory_order_release);
        const HRESULT hr = terminal.initializeEngine();
        check(FAILED(hr) && engineReleased(terminal) && ! terminal._engineInitializationError.empty(),
              std::format(L"initialization stage {} rolls back", stage));
    }

    {
        Terminal terminal;
        const HRESULT firstHr = terminal.initializeEngine();
        const HRESULT secondHr = terminal.initializeEngine();
        check(SUCCEEDED(firstHr) && secondHr == HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED) &&
                  terminal._runtime.IsLoaded() && terminal._ghosttyTerminal != nullptr,
              L"second initialization is refused without replacing the live engine");
    }

    {
        const TerminalPngDecodeThreadContext pngDecodeContext;
        check(SUCCEEDED(pngDecodeContext.Status()), L"PNG decoder context initializes on the selftest thread");
    }

    {
        Terminal terminal;
        const HRESULT windowHr = terminal.createWindow(GetDesktopWindow());
        const HRESULT initializeHr = SUCCEEDED(windowHr) ? terminal.initializeEngine() : windowHr;
        if (SUCCEEDED(initializeHr))
        {
            terminal._sessionGeneration.store(1u, std::memory_order_release);
            terminal._outputGeneration.store(0u, std::memory_order_release);
            terminal._outputConsumedGeneration.store(0u, std::memory_order_release);
            terminal._outputNotificationPending.store(false, std::memory_order_release);
            terminal.notifyOutputReady();
            MSG outputMessage{};
            const bool outputPosted = PeekMessageW(&outputMessage,
                                                   terminal._window.get(),
                                                   WndMsg::kTerminalOutputReady,
                                                   WndMsg::kTerminalOutputReady,
                                                   PM_REMOVE) != FALSE;
            const MessageRouteResult outputRoute = outputPosted
                ? terminal.routeSessionMessage(
                      terminal._window.get(), outputMessage.message, outputMessage.wParam, outputMessage.lParam)
                : MessageRouteResult{};
            check(outputPosted && outputRoute.handled &&
                      terminal._outputConsumedGeneration.load(std::memory_order_acquire) == 1u &&
                      ! terminal._outputNotificationPending.load(std::memory_order_acquire),
                  L"output-ready payload crosses the registry and clears the coalescing gate");

            NotifyKittyReady(&terminal);
            MSG kittyMessage{};
            const bool kittyPosted = PeekMessageW(&kittyMessage,
                                                  terminal._window.get(),
                                                  WndMsg::kTerminalKittyReady,
                                                  WndMsg::kTerminalKittyReady,
                                                  PM_REMOVE) != FALSE;
            const MessageRouteResult kittyRoute = kittyPosted
                ? terminal.routeSessionMessage(
                      terminal._window.get(), kittyMessage.message, kittyMessage.wParam, kittyMessage.lParam)
                : MessageRouteResult{};
            check(kittyPosted && kittyRoute.handled,
                  L"Kitty-ready payload crosses the registry on the current session generation");

            terminal._osc52Policy.store(Osc52Policy::Allow, std::memory_order_release);
            terminal._debugClipboardCallbackCount.store(0u, std::memory_order_release);
            terminal._debugClipboardReplyCount.store(0u, std::memory_order_release);
            terminal._debugClipboardReplyResult.store(
                GHOSTTY_CLIPBOARD_WRITE_RESULT_MAX_VALUE, std::memory_order_release);
            terminal._debugOsc52PayloadPostCount.store(0u, std::memory_order_release);
            constexpr std::string_view clipboardWrite = "\x1b]52;c;Y2xpcGJvYXJkLXRlc3Q=\x1b\\";
            terminal._runtime.terminalVtWrite(terminal._ghosttyTerminal,
                                              reinterpret_cast<const uint8_t*>(clipboardWrite.data()),
                                              clipboardWrite.size());
            check(terminal._debugClipboardCallbackCount.load(std::memory_order_acquire) == 1u &&
                      terminal._debugClipboardReplyCount.load(std::memory_order_acquire) == 1u &&
                      terminal._debugClipboardReplyResult.load(std::memory_order_acquire) ==
                          GHOSTTY_CLIPBOARD_WRITE_RESULT_BUSY &&
                      terminal._debugOsc52PayloadPostCount.load(std::memory_order_acquire) == 1u &&
                      terminal._osc52PostPending.load(std::memory_order_acquire),
                  L"admitted DLL invokes the production OSC 52 callback once and receives one BUSY reply");

            terminal._runtime.terminalVtWrite(terminal._ghosttyTerminal,
                                              reinterpret_cast<const uint8_t*>(clipboardWrite.data()),
                                              clipboardWrite.size());
            check(terminal._debugClipboardCallbackCount.load(std::memory_order_acquire) == 2u &&
                      terminal._debugClipboardReplyCount.load(std::memory_order_acquire) == 2u &&
                      terminal._debugClipboardReplyResult.load(std::memory_order_acquire) ==
                          GHOSTTY_CLIPBOARD_WRITE_RESULT_BUSY &&
                      terminal._debugOsc52PayloadPostCount.load(std::memory_order_acquire) == 1u,
                  L"pending production clipboard work remains coalesced without an inner runtime retry");

            NotifyKittyReady(&terminal);
            const size_t drainedPayloads = DrainPostedPayloadsForWindow(terminal._window.get());
            MSG staleKittyMessage{};
            const bool staleKittyQueued = PeekMessageW(&staleKittyMessage,
                                                       terminal._window.get(),
                                                       WndMsg::kTerminalKittyReady,
                                                       WndMsg::kTerminalKittyReady,
                                                       PM_REMOVE) != FALSE;
            const MessageRouteResult staleKittyRoute = staleKittyQueued
                ? terminal.routeSessionMessage(terminal._window.get(),
                                               staleKittyMessage.message,
                                               staleKittyMessage.wParam,
                                               staleKittyMessage.lParam)
                : MessageRouteResult{};
            terminal._outputNotificationPending.store(false, std::memory_order_release);
            terminal.notifyOutputReady();
            check(drainedPayloads >= 2u && staleKittyQueued && staleKittyRoute.handled &&
                      ! terminal._outputNotificationPending.load(std::memory_order_acquire),
                  L"teardown drain invalidates stale wakes and a closed registry fails posting without stranding the gate");
            InitPostedPayloadWindow(terminal._window.get());
        }
        else
        {
            check(false, std::format(L"production clipboard callback fixture initializes (hr=0x{:08X})",
                                     static_cast<uint32_t>(initializeHr)));
        }
    }

    return *failedTests == 0u ? S_OK : E_FAIL;
}
#endif

HRESULT Terminal::DebugRunCommandExperiencePerfSelfTests(
    unsigned int* passedTests,
    unsigned int* failedTests) noexcept
{
    if (passedTests == nullptr || failedTests == nullptr)
    {
        return E_POINTER;
    }
    *passedTests = 0u;
    *failedTests = 0u;
    const auto check = [passedTests, failedTests](bool condition) noexcept
    {
        condition ? ++*passedTests : ++*failedTests;
    };
    const auto percentile = [](std::vector<uint64_t> values, size_t numerator, size_t denominator)
    {
        std::ranges::sort(values);
        if (values.empty())
        {
            return uint64_t{0u};
        }
        const size_t index = std::min(values.size() - 1u, (values.size() * numerator + denominator - 1u) / denominator - 1u);
        return values[index];
    };
    const auto elapsedNanoseconds = [](std::chrono::steady_clock::time_point startedAt) noexcept
    {
        return static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::nanoseconds>(
            std::chrono::steady_clock::now() - startedAt).count());
    };
    const auto asMicroseconds = [](uint64_t nanoseconds) noexcept
    {
        return (nanoseconds + 999u) / 1'000u;
    };

    const HRESULT initializeHr = initializeEngine();
    check(SUCCEEDED(initializeHr));
    if (FAILED(initializeHr))
    {
        return E_FAIL;
    }

    std::string scrollbackFixture;
    for (size_t row = 0u; row < 64u; ++row)
    {
        scrollbackFixture.append(std::format("row {}\r\n", row));
    }
    _runtime.terminalVtWrite(
        _ghosttyTerminal, reinterpret_cast<const uint8_t*>(scrollbackFixture.data()), scrollbackFixture.size());
    scrollViewportTo(0u);
    GhosttyTerminalScrollbar enterScrollbar{};
    const bool enterStartedAboveBottom = getScrollbarState(enterScrollbar) && enterScrollbar.total > enterScrollbar.len && enterScrollbar.offset == 0u;
    const bool enterEncoded = encodeSpecialKey(VK_RETURN, 0, false);
    const bool enterFollowedBottom = getScrollbarState(enterScrollbar) &&
        enterScrollbar.offset == enterScrollbar.total - std::min(enterScrollbar.len, enterScrollbar.total);
    check(enterStartedAboveBottom && enterEncoded && enterFollowedBottom);

    int wheelRemainder = 0;
    intptr_t wheelRowChecksum = 0;
    const auto wheelStartedAt = std::chrono::steady_clock::now();
    for (size_t index = 0u; index < 1'000u; ++index)
    {
        const TerminalWheelScrollResult resolved = ResolveTerminalWheelScroll(
            wheelRemainder, index % 2u == 0u ? WHEEL_DELTA * 2 : -WHEEL_DELTA * 2, 3u, _rows);
        wheelRemainder = resolved.remainder;
        wheelRowChecksum += resolved.rows;
    }
    const uint64_t wheelDurationUs = Debug::Perf::ElapsedUs(wheelStartedAt);
    check(wheelRemainder == 0 && wheelRowChecksum == 0 && wheelDurationUs <= 50'000u);
    Debug::Perf::EmitDurationUs(L"terminal.wheel.resolve_batch_us", wheelDurationUs, 1'000u, 2'000u);

    constexpr size_t kRouteMessageCount = 50'000u;
    constexpr size_t kRouteBatchSize = 250u;
    std::vector<uint64_t> baselineDurations;
    std::vector<uint64_t> routeDurations;
    baselineDurations.reserve(kRouteMessageCount);
    routeDurations.reserve(kRouteMessageCount);
    uint64_t bypassChecksum = 0u;
    for (size_t batchStart = 0u; batchStart < kRouteMessageCount; batchStart += kRouteBatchSize)
    {
        const auto batchStartedAt = std::chrono::steady_clock::now();
        for (size_t index = 0u; index < kRouteBatchSize; ++index)
        {
            const auto startedAt = std::chrono::steady_clock::now();
            bypassChecksum ^= static_cast<uint64_t>(TerminalShortcutRoute::PassThrough) + batchStart + index;
            std::atomic_signal_fence(std::memory_order_seq_cst);
            baselineDurations.push_back(elapsedNanoseconds(startedAt));
        }
        Debug::Perf::EmitDurationUs(L"terminal.shortcut.bypass_baseline_batch_us",
                                    Debug::Perf::ElapsedUs(batchStartedAt),
                                    kRouteBatchSize,
                                    0u);
    }

    constexpr std::wstring_view kUnmappedCommand = L"cmd/app/terminalPerfUnmapped";
    constexpr std::wstring_view kPassThroughCommand = L"cmd/terminal/find";
    TerminalShortcutRequest request{};
    request.sizeBytes = sizeof(request);
    request.message = WM_KEYDOWN;
    request.virtualKey = VK_F24;
    TerminalShortcutRoute route = TerminalShortcutRoute::Blocked;
    for (size_t index = 0u; index < 1'000u; ++index)
    {
        const std::wstring_view command = (index & 1u) == 0u ? kUnmappedCommand : kPassThroughCommand;
        request.commandId = {command.data(), static_cast<uint32_t>(command.size())};
        static_cast<void>(RouteShortcut(&request, &route));
    }
    flushShortcutRouteMetrics();

    bool routeResultsValid = true;
    for (size_t batchStart = 0u; batchStart < kRouteMessageCount; batchStart += kRouteBatchSize)
    {
        const auto batchStartedAt = std::chrono::steady_clock::now();
        for (size_t index = 0u; index < kRouteBatchSize; ++index)
        {
            const size_t messageIndex = batchStart + index;
            const std::wstring_view command = (messageIndex & 1u) == 0u ? kUnmappedCommand : kPassThroughCommand;
            const TerminalShortcutRoute expected = (messageIndex & 1u) == 0u
                ? TerminalShortcutRoute::InvokeHostCommand
                : TerminalShortcutRoute::PassThrough;
            request.commandId = {command.data(), static_cast<uint32_t>(command.size())};
            const auto startedAt = std::chrono::steady_clock::now();
            const HRESULT routeHr = RouteShortcut(&request, &route);
            routeDurations.push_back(elapsedNanoseconds(startedAt));
            routeResultsValid = routeResultsValid && routeHr == S_OK && route == expected;
        }
        Debug::Perf::EmitDurationUs(L"terminal.shortcut.route_candidate_batch_us",
                                    Debug::Perf::ElapsedUs(batchStartedAt),
                                    kRouteBatchSize,
                                    0u);
    }
    flushShortcutRouteMetrics();
    const uint64_t baselineMedianNs = percentile(baselineDurations, 50u, 100u);
    const uint64_t baselineP95Ns = percentile(baselineDurations, 95u, 100u);
    const uint64_t routeMedianNs = percentile(routeDurations, 50u, 100u);
    const uint64_t routeP95Ns = percentile(routeDurations, 95u, 100u);
    const uint64_t allowedRouteP95Ns = baselineP95Ns + baselineP95Ns * 15u / 100u + 2'000u;
    check(routeResultsValid && bypassChecksum != (std::numeric_limits<uint64_t>::max)());
    check(routeP95Ns <= 25'000u && routeP95Ns <= allowedRouteP95Ns);
    Debug::Perf::EmitDurationUs(
        L"terminal.shortcut.route_baseline_median_us", asMicroseconds(baselineMedianNs), kRouteMessageCount, 0u);
    Debug::Perf::EmitDurationUs(
        L"terminal.shortcut.route_baseline_p95_us", asMicroseconds(baselineP95Ns), kRouteMessageCount, 0u);
    Debug::Perf::EmitDurationUs(
        L"terminal.shortcut.route_candidate_median_us", asMicroseconds(routeMedianNs), kRouteMessageCount, 0u);
    Debug::Perf::EmitDurationUs(
        L"terminal.shortcut.route_candidate_p95_us", asMicroseconds(routeP95Ns), kRouteMessageCount, 0u);

    constexpr std::wstring_view kScrollCommand = L"cmd/terminal/scroll/lineDown";
    request.commandId = {kScrollCommand.data(), static_cast<uint32_t>(kScrollCommand.size())};
    const uint64_t scrollMutationsBefore = _debugScrollViewportMutationCount;
    const auto scrollStartedAt = std::chrono::steady_clock::now();
    bool scrollResultsValid = true;
    for (size_t index = 0u; index < 1'000u; ++index)
    {
        scrollResultsValid = scrollResultsValid && RouteShortcut(&request, &route) == S_OK && route == TerminalShortcutRoute::Handled;
    }
    flushShortcutRouteMetrics();
    const uint64_t scrollDurationUs = Debug::Perf::ElapsedUs(scrollStartedAt);
    check(scrollResultsValid && _debugScrollViewportMutationCount - scrollMutationsBefore == 1'000u);
    check(scrollDurationUs <= 50'000u);
    Debug::Perf::EmitDurationUs(L"terminal.shortcut.scroll_repeat_batch_us", scrollDurationUs, 1'000u, 1'000u);

    constexpr std::array<std::wstring_view, 3u> kZoomCommands{{
        L"cmd/terminal/font/increase", L"cmd/terminal/font/decrease", L"cmd/terminal/font/reset"}};
    const auto zoomStartedAt = std::chrono::steady_clock::now();
    bool zoomResultsValid = true;
    for (size_t cycle = 0u; cycle < 100u; ++cycle)
    {
        for (const std::wstring_view command : kZoomCommands)
        {
            request.commandId = {command.data(), static_cast<uint32_t>(command.size())};
            zoomResultsValid = zoomResultsValid && RouteShortcut(&request, &route) == S_OK && route == TerminalShortcutRoute::Handled;
        }
    }
    flushShortcutRouteMetrics();
    check(zoomResultsValid && std::abs(_liveFontSizeDip - _config.fontSizeDip) < 0.01f && ! _renderTarget && ! _textFormat);
    Debug::Perf::EmitDurationUs(
        L"terminal.shortcut.zoom_cycle_batch_us", Debug::Perf::ElapsedUs(zoomStartedAt), 100u, 0u);

    constexpr size_t kFindLineCount = 100'000u;
    std::wstring findSnapshot;
    findSnapshot.reserve(kFindLineCount * 48u);
    for (size_t line = 0u; line < kFindLineCount; ++line)
    {
        if (line != 0u)
        {
            findSnapshot.push_back(L'\n');
        }
        findSnapshot.append(std::format(L"line {0:06} terminal needle alpha beta gamma", line));
    }
    const auto snapshotPublishStartedAt = std::chrono::steady_clock::now();
    const std::wstring_view publishedSnapshot(findSnapshot);
    const uint64_t snapshotPublishUs = Debug::Perf::ElapsedUs(snapshotPublishStartedAt);
    check(snapshotPublishUs < 4'000u && publishedSnapshot.size() == findSnapshot.size());
    Debug::Perf::EmitDurationUs(
        L"terminal.find.snapshot_batch_us", snapshotPublishUs, findSnapshot.size() * sizeof(wchar_t), kFindLineCount);

    constexpr std::array<std::wstring_view, 5u> kFindQueries{{
        L"line", L"TERMINAL", L"needle", L"alpha beta", L"not-present"}};
    std::vector<uint64_t> findDurations;
    findDurations.reserve(1'000u);
    bool findResultsValid = true;
    for (size_t index = 0u; index < 1'000u; ++index)
    {
        const auto startedAt = std::chrono::steady_clock::now();
        const TerminalCommandSurfaceModel::FindResult result = TerminalCommandSurfaceModel::Find(
            publishedSnapshot, kFindQueries[index % kFindQueries.size()], 128u);
        const uint64_t durationUs = Debug::Perf::ElapsedUs(startedAt);
        findDurations.push_back(durationUs);
        findResultsValid = findResultsValid && result.scannedLines == kFindLineCount && ! result.cancelled;
        Debug::Perf::EmitDurationUs(
            L"terminal.find.query_to_visible_us", durationUs, result.scannedLines, result.totalMatches);
    }
    std::stop_source cancelledFind;
    cancelledFind.request_stop();
    const bool findCancellation = TerminalCommandSurfaceModel::Find(
        publishedSnapshot, L"needle", 128u, cancelledFind.get_token()).cancelled;
    const uint64_t findP95Us = percentile(findDurations, 95u, 100u);
    check(findResultsValid && findCancellation && findP95Us <= 50'000u);
    Debug::Perf::EmitValue(L"terminal.find.cancelled_count", findCancellation ? 1u : 0u);

    constexpr size_t kPersistedHistoryCount = 9'000u;
    constexpr size_t kSessionHistoryCount = 1'000u;
    constexpr size_t kHistoryEntryCharacters = 200u;
    const auto suggestionLoadStartedAt = std::chrono::steady_clock::now();
    std::vector<std::wstring> persistedHistory;
    std::vector<std::wstring> currentHistory;
    persistedHistory.reserve(kPersistedHistoryCount);
    currentHistory.reserve(kSessionHistoryCount);
    for (size_t index = 0u; index < kPersistedHistoryCount + kSessionHistoryCount; ++index)
    {
        std::wstring entry = std::format(L"cmd {0:05} ", index);
        entry.append(kHistoryEntryCharacters - entry.size(), static_cast<wchar_t>(L'a' + index % 26u));
        (index < kPersistedHistoryCount ? persistedHistory : currentHistory).push_back(std::move(entry));
    }
    const uint64_t suggestionLoadDurationUs = Debug::Perf::ElapsedUs(suggestionLoadStartedAt);
    check(suggestionLoadDurationUs <= 100'000u);
    Debug::Perf::EmitDurationUs(L"terminal.suggestions.load_batch_us",
                                suggestionLoadDurationUs,
                                kPersistedHistoryCount + kSessionHistoryCount,
                                (kPersistedHistoryCount + kSessionHistoryCount) * kHistoryEntryCharacters * sizeof(wchar_t));

    std::vector<uint64_t> suggestionOpenDurations;
    suggestionOpenDurations.reserve(100u);
    bool suggestionResultsValid = true;
    for (size_t iteration = 0u; iteration < 100u; ++iteration)
    {
        const auto startedAt = std::chrono::steady_clock::now();
        const TerminalCommandSurfaceModel::SuggestionResult result = TerminalCommandSurfaceModel::FilterSuggestions(
            persistedHistory, currentHistory, L"", 100u);
        const uint64_t durationUs = Debug::Perf::ElapsedUs(startedAt);
        suggestionOpenDurations.push_back(durationUs);
        suggestionResultsValid = suggestionResultsValid && result.sourceEntries == 10'000u &&
            result.sourceBytes == 4'000'000u && result.rows.size() == 100u && ! result.cancelled;
        Debug::Perf::EmitDurationUs(
            L"terminal.suggestions.open_to_visible_us", durationUs, result.sourceEntries, result.sourceBytes);
    }

    std::vector<uint64_t> suggestionFilterDurations;
    suggestionFilterDurations.reserve(1'000u);
    for (size_t iteration = 0u; iteration < 1'000u; ++iteration)
    {
        const std::wstring query = std::format(L"cmd {0:05}", iteration % 10'000u);
        const auto startedAt = std::chrono::steady_clock::now();
        const TerminalCommandSurfaceModel::SuggestionResult result = TerminalCommandSurfaceModel::FilterSuggestions(
            persistedHistory, currentHistory, query, 100u);
        const uint64_t durationUs = Debug::Perf::ElapsedUs(startedAt);
        suggestionFilterDurations.push_back(durationUs);
        suggestionResultsValid = suggestionResultsValid && result.sourceEntries == 10'000u &&
            result.sourceBytes == 4'000'000u && result.rows.size() == 1u && ! result.cancelled;
        Debug::Perf::EmitDurationUs(
            L"terminal.suggestions.filter_to_visible_us", durationUs, result.sourceEntries, result.sourceBytes);
    }
    std::stop_source cancelledSuggestions;
    cancelledSuggestions.request_stop();
    const bool suggestionsCancelled = TerminalCommandSurfaceModel::FilterSuggestions(
        persistedHistory, currentHistory, L"cmd", 100u, cancelledSuggestions.get_token()).cancelled;
    check(suggestionResultsValid && suggestionsCancelled &&
          percentile(suggestionOpenDurations, 95u, 100u) <= 100'000u &&
          percentile(suggestionFilterDurations, 95u, 100u) <= 16'700u);
    Debug::Perf::EmitValue(L"terminal.suggestions.cancelled_count", suggestionsCancelled ? 1u : 0u);

    return *failedTests == 0u ? S_OK : E_FAIL;
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderTerminalDebugCommandExperiencePerfSelfTests(
    unsigned int* passedTests,
    unsigned int* failedTests) noexcept
{
    Terminal terminal;
    return terminal.DebugRunCommandExperiencePerfSelfTests(passedTests, failedTests);
}

HRESULT Terminal::DebugRunKittyCacheSelfTests(
    unsigned int* passedTests,
    unsigned int* failedTests) noexcept
{
    if (passedTests == nullptr || failedTests == nullptr)
    {
        return E_POINTER;
    }
    const auto check = [passedTests, failedTests](bool condition) noexcept
    {
        if (condition)
        {
            ++*passedTests;
        }
        else
        {
            ++*failedTests;
        }
    };

    const HRESULT initializeHr = initializeEngine();
    check(SUCCEEDED(initializeHr));
    if (FAILED(initializeHr))
    {
        return initializeHr;
    }

    KittyImageSnapshot queryState{};
    queryState.state = KittyImageState::ConversionPending;
    queryState.pendingRequestId = 41u;
    const bool enginePendingHandled = handleKittyImageQueryResult(queryState, GHOSTTY_NO_VALUE);
    check(enginePendingHandled && queryState.state == KittyImageState::EnginePending &&
          queryState.pendingRequestId == 0u);
    check(needsKittyReplacementWork(queryState, 41u, false));
    resetKittyPendingForStorageChange(queryState);
    check(queryState.state == KittyImageState::Missing && queryState.pendingRequestId == 0u);

    queryState.state = KittyImageState::ConversionPending;
    queryState.pendingRequestId = 42u;
    const bool successHandled = handleKittyImageQueryResult(queryState, GHOSTTY_SUCCESS);
    const bool currentResult = isCurrentKittyConversionResult(queryState, 42u);
    const bool staleResult = isCurrentKittyConversionResult(queryState, 41u);
    check(! successHandled && currentResult && ! staleResult);

    queryState.state = KittyImageState::Ready;
    queryState.pendingRequestId = 0u;
    const bool failureHandled = handleKittyImageQueryResult(queryState, GHOSTTY_INVALID_VALUE);
    check(failureHandled && queryState.state == KittyImageState::Rejected && queryState.pendingRequestId == 0u);
    check(! needsKittyReplacementWork(queryState, 42u, false));
    resetKittyPendingForStorageChange(queryState);
    check(queryState.state == KittyImageState::Missing &&
          needsKittyReplacementWork(queryState, 42u, false));

    const bool terminalSized =
        _runtime.terminalResize(_ghosttyTerminal, _columns, _rows, 8u, 16u) == GHOSTTY_SUCCESS;

    constexpr std::string_view imageA =
        "\x1b_Ga=T,t=d,f=24,i=101,p=101,s=1,v=1,c=1,r=1,q=2;/wAA\x1b\\";
    constexpr std::string_view scrollback =
        "\r\n01\r\n02\r\n03\r\n04\r\n05\r\n06\r\n07\r\n08\r\n09\r\n10"
        "\r\n11\r\n12\r\n13\r\n14\r\n15\r\n16\r\n17\r\n18\r\n19\r\n20"
        "\r\n21\r\n22\r\n23\r\n24\r\n25\r\n26\r\n27\r\n28\r\n29\r\n30"
        "\r\n31\r\n32\r\n33\r\n34\r\n35\r\n36\r\n37\r\n38\r\n39\r\n40\r\n";
    constexpr std::string_view imageB =
        "\x1b_Ga=T,t=d,f=24,i=202,p=202,s=1,v=1,c=1,r=1,q=2;AP8A\x1b\\";
    _runtime.terminalVtWrite(_ghosttyTerminal, reinterpret_cast<const uint8_t*>(imageA.data()), imageA.size());
    _runtime.terminalVtWrite(
        _ghosttyTerminal, reinterpret_cast<const uint8_t*>(scrollback.data()), scrollback.size());
    _runtime.terminalVtWrite(_ghosttyTerminal, reinterpret_cast<const uint8_t*>(imageB.data()), imageB.size());

    GhosttyTerminalScrollbar scrollbar{};
    const bool scrollbarReady = terminalSized &&
        _runtime.terminalGet(_ghosttyTerminal, GHOSTTY_TERMINAL_DATA_SCROLLBAR, &scrollbar) == GHOSTTY_SUCCESS &&
        scrollbar.total > scrollbar.len;
    GhosttyTerminalScrollViewport scroll{};
    scroll.tag = GHOSTTY_SCROLL_VIEWPORT_ROW;
    scroll.value.row = 0u;
    if (scrollbarReady)
    {
        _runtime.terminalScrollViewport(_ghosttyTerminal, scroll);
    }

    TerminalKittyGenerationWork firstWork;
    std::vector<KittyPlacementSnapshot> firstPlacements;
    bool firstSubmitted = false;
    bool firstCaptured = false;
    {
        std::scoped_lock lock(_terminalMutex);
        firstCaptured = captureKittyGraphicsLocked(firstWork, firstPlacements, firstSubmitted);
    }
    const bool firstImageCaptured = std::ranges::any_of(
        firstWork.images,
        [](const TerminalKittyImageWork& image) noexcept { return image.imageId == 101u; });
    const bool firstPlacementVisible = std::ranges::any_of(
        firstPlacements,
        [](const KittyPlacementSnapshot& placement) noexcept { return placement.imageKey.imageId == 101u; });
    const uint64_t firstProbe = (scrollbarReady ? 1u : 0u) | (firstCaptured ? 2u : 0u) |
        (firstSubmitted ? 4u : 0u) | (firstWork.storageGeneration != 0u ? 8u : 0u) |
        (firstImageCaptured ? 16u : 0u) | (firstPlacementVisible ? 32u : 0u);
    Debug::Perf::EmitValue(L"terminal.kitty.visible_delta_first_probe", firstProbe);
    check(scrollbarReady && firstCaptured && firstSubmitted && firstWork.storageGeneration != 0u &&
          firstImageCaptured && firstPlacementVisible);
    if (! firstCaptured || firstWork.storageGeneration == 0u)
    {
        return E_FAIL;
    }

    const uint64_t firstStorageGeneration = firstWork.storageGeneration;
    const uint64_t firstRequestId = firstWork.requestId;
    const uint64_t imageAGeneration = firstWork.images.empty() ? 0u : firstWork.images.front().imageGeneration;
    const bool firstSubmitAccepted = firstSubmitted && _kittyPipeline.Submit(std::move(firstWork));
    const auto waitForReady = [this](const KittyImageKey& imageKey, uint64_t requestId) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + std::chrono::seconds(2);
        while (std::chrono::steady_clock::now() < deadline)
        {
            consumeKittyReady();
            const auto imageIterator = _kittyImages.find(imageKey);
            if (imageIterator != _kittyImages.end() && imageIterator->second.state == KittyImageState::Ready &&
                imageIterator->second.pendingRequestId == 0u)
            {
                return true;
            }
            if (_kittyPipeline.GetStats().latestRequestId > requestId)
            {
                return false;
            }
            std::this_thread::yield();
        }
        return false;
    };
    const KittyImageKey imageAKey{101u, imageAGeneration};
    const bool firstReady = firstSubmitAccepted && waitForReady(imageAKey, firstRequestId);
    check(firstSubmitAccepted && firstReady);

    scroll.value.row = static_cast<size_t>(scrollbar.total - scrollbar.len);
    _runtime.terminalScrollViewport(_ghosttyTerminal, scroll);

    TerminalKittyGenerationWork revealedWork;
    std::vector<KittyPlacementSnapshot> revealedPlacements;
    bool revealedSubmitted = false;
    bool revealedCaptured = false;
    {
        std::scoped_lock lock(_terminalMutex);
        revealedCaptured = captureKittyGraphicsLocked(revealedWork, revealedPlacements, revealedSubmitted);
    }
    const bool revealedImageCopied = std::ranges::any_of(
        revealedWork.images,
        [](const TerminalKittyImageWork& image) noexcept { return image.imageId == 202u; });
    const bool revealedPlacementVisible = std::ranges::any_of(
        revealedPlacements,
        [](const KittyPlacementSnapshot& placement) noexcept { return placement.imageKey.imageId == 202u; });
    const uint64_t revealedProbe = (revealedCaptured ? 1u : 0u) |
        (revealedWork.storageGeneration == firstStorageGeneration ? 2u : 0u) |
        (revealedSubmitted ? 4u : 0u) | (revealedImageCopied ? 8u : 0u) |
        (revealedPlacementVisible ? 16u : 0u);
    Debug::Perf::EmitValue(L"terminal.kitty.visible_delta_revealed_probe", revealedProbe);
    check(revealedCaptured && revealedWork.storageGeneration == firstStorageGeneration && revealedSubmitted &&
          revealedImageCopied && revealedPlacementVisible);

    const uint64_t revealedRequestId = revealedWork.requestId;
    const uint64_t imageBGeneration = revealedWork.images.empty() ? 0u : revealedWork.images.front().imageGeneration;
    const bool revealedSubmitAccepted = revealedSubmitted && _kittyPipeline.Submit(std::move(revealedWork));
    const KittyImageKey imageBKey{202u, imageBGeneration};
    const bool revealedReady = revealedSubmitAccepted && waitForReady(imageBKey, revealedRequestId);
    const auto retainedA = _kittyImages.find(imageAKey);
    check(revealedReady && retainedA != _kittyImages.end() && retainedA->second.state == KittyImageState::Ready);

    constexpr std::string_view imageAReplacement =
        "\x1b_Ga=T,t=d,f=24,i=101,p=101,s=1,v=1,c=1,r=1,q=2;AAH/\x1b\\";
    const size_t supersededBeforeReplacement = _kittySupersededImageCount;
    _runtime.terminalVtWrite(
        _ghosttyTerminal,
        reinterpret_cast<const uint8_t*>(imageAReplacement.data()),
        imageAReplacement.size());
    TerminalKittyGenerationWork replacementWork;
    std::vector<KittyPlacementSnapshot> replacementPlacements;
    bool replacementSubmitted = false;
    bool replacementCaptured = false;
    {
        std::scoped_lock lock(_terminalMutex);
        replacementCaptured = captureKittyGraphicsLocked(
            replacementWork, replacementPlacements, replacementSubmitted);
    }
    const auto replacementDeadline = std::chrono::steady_clock::now() + std::chrono::seconds(2);
    while (replacementCaptured && ! replacementSubmitted &&
           std::chrono::steady_clock::now() < replacementDeadline)
    {
        std::this_thread::yield();
        std::scoped_lock lock(_terminalMutex);
        replacementCaptured = captureKittyGraphicsLocked(
            replacementWork, replacementPlacements, replacementSubmitted);
    }
    const auto replacementImage = std::ranges::find_if(
        replacementWork.images,
        [](const TerminalKittyImageWork& image) noexcept { return image.imageId == 101u; });
    const uint64_t replacementImageGeneration = replacementImage == replacementWork.images.end()
        ? 0u
        : replacementImage->imageGeneration;
    const KittyImageKey replacementImageKey{101u, replacementImageGeneration};
    const bool replacementPlacementVisible = std::ranges::any_of(
        replacementPlacements,
        [&replacementImageKey](const KittyPlacementSnapshot& placement) noexcept
        { return placement.imageKey == replacementImageKey; });
    const bool oldGenerationRemoved = _kittyImages.find(imageAKey) == _kittyImages.end();
    check(replacementCaptured && replacementSubmitted && replacementPlacementVisible &&
          replacementImageGeneration != 0u && replacementImageGeneration != imageAGeneration &&
          oldGenerationRemoved && _kittySupersededImageCount == supersededBeforeReplacement + 1u);
    const uint64_t replacementRequestId = replacementWork.requestId;
    const uint64_t activeStorageGeneration = replacementWork.storageGeneration;
    const bool replacementSubmitAccepted = replacementSubmitted && _kittyPipeline.Submit(std::move(replacementWork));
    const bool replacementReady =
        replacementSubmitAccepted && waitForReady(replacementImageKey, replacementRequestId);
    const auto retainedB = _kittyImages.find(imageBKey);
    check(replacementSubmitAccepted && replacementReady && retainedB != _kittyImages.end() &&
          retainedB->second.state == KittyImageState::Ready);

    const bool resized = _runtime.terminalResize(_ghosttyTerminal, 100u, 30u, 8u, 16u) == GHOSTTY_SUCCESS;
    TerminalKittyGenerationWork resizedWork;
    std::vector<KittyPlacementSnapshot> resizedPlacements;
    bool resizedSubmitted = false;
    bool resizedCaptured = false;
    {
        std::scoped_lock lock(_terminalMutex);
        resizedCaptured = captureKittyGraphicsLocked(resizedWork, resizedPlacements, resizedSubmitted);
    }
    check(resized && resizedCaptured && resizedWork.storageGeneration == activeStorageGeneration && ! resizedSubmitted);

    constexpr std::string_view deleteAllImages = "\x1b_Ga=d,d=A,q=2;\x1b\\";
    _runtime.terminalVtWrite(
        _ghosttyTerminal, reinterpret_cast<const uint8_t*>(deleteAllImages.data()), deleteAllImages.size());
    TerminalKittyGenerationWork deletedWork;
    std::vector<KittyPlacementSnapshot> deletedPlacements;
    bool deletedSubmitted = false;
    bool deletedCaptured = false;
    {
        std::scoped_lock lock(_terminalMutex);
        deletedCaptured = captureKittyGraphicsLocked(deletedWork, deletedPlacements, deletedSubmitted);
    }
    check(deletedCaptured && ! deletedSubmitted && deletedPlacements.empty() && _kittyImages.empty() &&
          _kittyCachedBytes == 0u);

    constexpr uint32_t highCardinalityRepeatedImageId = 303u;
    constexpr uint32_t highCardinalityFirstUniqueImageId = 10'000u;
    constexpr uint32_t highCardinalityUniqueImageCount = 1'024u;
    constexpr uint32_t highCardinalityRepeatedVisiblePlacementCount = 1'024u;
    constexpr std::string_view highCardinalityImage =
        "\x1b_Ga=t,t=d,f=24,i=303,s=1,v=1,q=2;000A\x1b\\";
    _runtime.terminalVtWrite(
        _ghosttyTerminal,
        reinterpret_cast<const uint8_t*>(highCardinalityImage.data()),
        highCardinalityImage.size());
    for (uint32_t imageOffset = 0u; imageOffset < highCardinalityUniqueImageCount; ++imageOffset)
    {
        const std::string transmit = std::format("\x1b_Ga=t,t=d,f=24,i={},s=1,v=1,q=2;000A\x1b\\",
                                                 highCardinalityFirstUniqueImageId + imageOffset);
        _runtime.terminalVtWrite(
            _ghosttyTerminal, reinterpret_cast<const uint8_t*>(transmit.data()), transmit.size());
    }
    for (uint32_t placementId = 1u; placementId <= 1'024u; ++placementId)
    {
        const std::string display = std::format("\x1b_Ga=p,i=303,p={},c=1,r=1,q=2;\x1b\\", placementId);
        _runtime.terminalVtWrite(
            _ghosttyTerminal, reinterpret_cast<const uint8_t*>(display.data()), display.size());
    }
    _runtime.terminalVtWrite(
        _ghosttyTerminal, reinterpret_cast<const uint8_t*>(scrollback.data()), scrollback.size());
    for (uint32_t placementId = 1'025u; placementId <= 2'048u; ++placementId)
    {
        const std::string display =
            std::format("\x1b_Ga=p,i=303,p={},U=1,c=1,r=1,q=2;\x1b\\", placementId);
        _runtime.terminalVtWrite(
            _ghosttyTerminal, reinterpret_cast<const uint8_t*>(display.data()), display.size());
    }
    for (uint32_t imageOffset = 0u; imageOffset < highCardinalityUniqueImageCount; ++imageOffset)
    {
        const std::string display = std::format("\x1b_Ga=p,i={},p={},c=1,r=1,q=2;\x1b\\",
                                                highCardinalityFirstUniqueImageId + imageOffset,
                                                2'049u + imageOffset);
        _runtime.terminalVtWrite(
            _ghosttyTerminal, reinterpret_cast<const uint8_t*>(display.data()), display.size());
    }
    for (uint32_t placementId = 3'073u; placementId <= kMaximumKittyPlacements; ++placementId)
    {
        const std::string display =
            std::format("\x1b_Ga=p,i={},p={},c=1,r=1,q=2;\x1b\\", highCardinalityRepeatedImageId, placementId);
        _runtime.terminalVtWrite(
            _ghosttyTerminal, reinterpret_cast<const uint8_t*>(display.data()), display.size());
    }
    constexpr std::string_view rejectedPlacement = "\x1b_Ga=p,i=303,p=4097,c=1,r=1,q=2;\x1b\\";
    _runtime.terminalVtWrite(
        _ghosttyTerminal,
        reinterpret_cast<const uint8_t*>(rejectedPlacement.data()),
        rejectedPlacement.size());

    TerminalKittyGenerationWork highCardinalityWork;
    std::vector<KittyPlacementSnapshot> highCardinalityPlacements;
    bool highCardinalitySubmitted = false;
    bool highCardinalityCaptured = false;
    const auto highCardinalityCaptureStartedAt = std::chrono::steady_clock::now();
    {
        std::scoped_lock lock(_terminalMutex);
        highCardinalityCaptured = captureKittyGraphicsLocked(
            highCardinalityWork, highCardinalityPlacements, highCardinalitySubmitted);
    }
    const uint64_t highCardinalityCaptureUs = Debug::Perf::ElapsedUs(highCardinalityCaptureStartedAt);
    Debug::Perf::EmitDurationUs(L"terminal.kitty.capture_lock_us",
                                highCardinalityCaptureUs,
                                _kittyCaptureSourceBytes,
                                _kittyCaptureTotalPlacements,
                                highCardinalityCaptured ? S_OK : E_FAIL);
    Debug::Perf::EmitValue(L"terminal.kitty.placement_total_count", _kittyCaptureTotalPlacements);
    Debug::Perf::EmitValue(L"terminal.kitty.placement_visible_count", _kittyCaptureVisiblePlacements);
    Debug::Perf::EmitValue(L"terminal.kitty.placement_offscreen_count", _kittyCaptureOffscreenPlacements);
    Debug::Perf::EmitValue(L"terminal.kitty.placement_virtual_count", _kittyCaptureVirtualPlacements);
    Debug::Perf::EmitValue(L"terminal.kitty.visible_image_key_count", _kittyCaptureVisibleImageKeys);
    Debug::Perf::EmitValue(L"terminal.kitty.engine_placement_count", _kittyCaptureEnginePlacementCount);
    Debug::Perf::EmitValue(L"terminal.kitty.engine_placement_bytes", _kittyCaptureEnginePlacementBytes);
    Debug::Perf::EmitValue(L"terminal.kitty.rejected_placement_count", 1u);
    std::vector<KittyImageKey> legacyImageKeys;
    legacyImageKeys.reserve(highCardinalityWork.images.size());
    uint64_t legacyCaptureComparisons = 0u;
    for (const KittyPlacementSnapshot& placement : highCardinalityPlacements)
    {
        bool found = false;
        for (const KittyImageKey& imageKey : legacyImageKeys)
        {
            ++legacyCaptureComparisons;
            if (imageKey == placement.imageKey)
            {
                found = true;
                break;
            }
        }
        if (! found)
        {
            legacyImageKeys.push_back(placement.imageKey);
        }
    }
    constexpr int32_t belowBackgroundMaximum = (std::numeric_limits<int32_t>::min)() / 2;
    uint64_t legacyDrawComparisons = 0u;
    const auto legacyDrawStartedAt = std::chrono::steady_clock::now();
    for (const GhosttyKittyPlacementLayer layer : {GHOSTTY_KITTY_PLACEMENT_LAYER_BELOW_BG,
                                                   GHOSTTY_KITTY_PLACEMENT_LAYER_BELOW_TEXT,
                                                   GHOSTTY_KITTY_PLACEMENT_LAYER_ABOVE_TEXT})
    {
        for (const KittyPlacementSnapshot& placement : highCardinalityPlacements)
        {
            const bool matches = layer == GHOSTTY_KITTY_PLACEMENT_LAYER_BELOW_BG
                    ? placement.z < belowBackgroundMaximum
                    : layer == GHOSTTY_KITTY_PLACEMENT_LAYER_BELOW_TEXT
                    ? placement.z >= belowBackgroundMaximum && placement.z < 0
                    : layer == GHOSTTY_KITTY_PLACEMENT_LAYER_ABOVE_TEXT && placement.z >= 0;
            if (! matches)
            {
                continue;
            }
            for (const KittyImageKey& imageKey : legacyImageKeys)
            {
                ++legacyDrawComparisons;
                if (imageKey == placement.imageKey)
                {
                    break;
                }
            }
        }
    }
    const uint64_t legacyDrawUs = Debug::Perf::ElapsedUs(legacyDrawStartedAt);
    Debug::Perf::EmitValue(L"terminal.kitty.legacy_linear_capture_comparison_count", legacyCaptureComparisons);
    Debug::Perf::EmitValue(L"terminal.kitty.legacy_linear_draw_comparison_count", legacyDrawComparisons);
    Debug::Perf::EmitDurationUs(
        L"terminal.kitty.legacy_linear_draw_lookup_us", std::max<uint64_t>(1u, legacyDrawUs), legacyDrawComparisons);
    const uint64_t visibleImageKeyCount = highCardinalityUniqueImageCount + 1u;
    const uint64_t minimumLegacyCaptureComparisons =
        visibleImageKeyCount * (visibleImageKeyCount - 1u) / 2u +
        highCardinalityRepeatedVisiblePlacementCount - 1u;
    const uint64_t maximumLegacyCaptureComparisons =
        visibleImageKeyCount * (visibleImageKeyCount - 1u) / 2u +
        static_cast<uint64_t>(highCardinalityRepeatedVisiblePlacementCount - 1u) * visibleImageKeyCount;
    const size_t highCardinalitySourceBytes = std::ranges::fold_left(
        highCardinalityWork.images, size_t{0u}, [](size_t total, const TerminalKittyImageWork& image) noexcept
        { return total + image.source.size(); });
    const bool highCardinalityShapeReady = highCardinalityCaptured && highCardinalitySubmitted &&
        _kittyCaptureTotalPlacements == kMaximumKittyPlacements &&
        _kittyCaptureEnginePlacementCount == kMaximumKittyPlacements &&
        _kittyCaptureEnginePlacementBytes > 0u &&
        _kittyCaptureEnginePlacementBytes <= kMaximumKittyPlacementBytes &&
        _kittyCaptureOffscreenPlacements >= 1'024u && _kittyCaptureVirtualPlacements == 1'024u &&
        _kittyCaptureVisiblePlacements + _kittyCaptureOffscreenPlacements + _kittyCaptureVirtualPlacements ==
            _kittyCaptureTotalPlacements &&
        _kittyCaptureVisibleImageKeys == highCardinalityUniqueImageCount + 1u &&
        highCardinalityPlacements.size() == _kittyCaptureVisiblePlacements &&
        highCardinalityWork.images.size() == highCardinalityUniqueImageCount + 1u &&
        highCardinalitySourceBytes == (highCardinalityUniqueImageCount + 1u) * 3u &&
        legacyImageKeys.size() == visibleImageKeyCount &&
        legacyCaptureComparisons >= minimumLegacyCaptureComparisons &&
        legacyCaptureComparisons <= maximumLegacyCaptureComparisons &&
        legacyDrawComparisons == legacyCaptureComparisons + visibleImageKeyCount &&
        legacyDrawComparisons > static_cast<uint64_t>(_kittyCaptureVisiblePlacements) * 250u;
    check(highCardinalityShapeReady);

    const uint64_t highCardinalityRequestId = highCardinalityWork.requestId;
    const auto highCardinalityRepeatedImage = std::ranges::find_if(
        highCardinalityWork.images,
        [](const TerminalKittyImageWork& image) noexcept
        { return image.imageId == highCardinalityRepeatedImageId; });
    const uint64_t highCardinalityImageGeneration = highCardinalityRepeatedImage == highCardinalityWork.images.end()
        ? 0u
        : highCardinalityRepeatedImage->imageGeneration;
    const bool highCardinalitySubmitAccepted =
        highCardinalitySubmitted && _kittyPipeline.Submit(std::move(highCardinalityWork));
    const KittyImageKey highCardinalityImageKey{highCardinalityRepeatedImageId, highCardinalityImageGeneration};
    const bool highCardinalityReady = highCardinalitySubmitAccepted &&
        waitForReady(highCardinalityImageKey, highCardinalityRequestId);
    TerminalKittyGenerationWork cacheReuseWork;
    std::vector<KittyPlacementSnapshot> cacheReusePlacements;
    bool cacheReuseSubmitted = false;
    bool cacheReuseCaptured = false;
    const auto cacheReuseCaptureStartedAt = std::chrono::steady_clock::now();
    if (highCardinalityReady)
    {
        std::scoped_lock lock(_terminalMutex);
        cacheReuseCaptured = captureKittyGraphicsLocked(cacheReuseWork, cacheReusePlacements, cacheReuseSubmitted);
    }
    const uint64_t cacheReuseCaptureUs = Debug::Perf::ElapsedUs(cacheReuseCaptureStartedAt);
    Debug::Perf::EmitDurationUs(L"terminal.kitty.cache_reuse_capture_lock_us",
                                std::max<uint64_t>(1u, cacheReuseCaptureUs),
                                _kittyCaptureSourceBytes,
                                _kittyCaptureTotalPlacements,
                                cacheReuseCaptured ? S_OK : E_FAIL);
    Debug::Perf::EmitValue(L"terminal.kitty.source_bytes", highCardinalitySourceBytes);
    Debug::Perf::EmitValue(L"terminal.kitty.cache_bytes", _kittyCachedBytes);
    Debug::Perf::EmitValue(L"terminal.kitty.pinned_bytes", _kittyCapturePinnedBytes);
    Debug::Perf::EmitValue(L"terminal.kitty.evicted_bytes", _kittyEvictedBytes);
    Debug::Perf::EmitValue(L"terminal.kitty.evicted_image_count", _kittyEvictedImageCount);
    Debug::Perf::EmitValue(L"terminal.kitty.missing_image_key_count", _kittyCaptureMissingImageKeys);
    Debug::Perf::EmitValue(L"terminal.kitty.deferred_image_key_count", _kittyCaptureDeferredImageKeys);
    Debug::Perf::EmitValue(L"terminal.kitty.rejected_image_key_count", _kittyCaptureRejectedImageKeys);
    _kittyPlacements = cacheReuseCaptured ? std::move(cacheReusePlacements) : std::move(highCardinalityPlacements);
    _kittyFrameImageLookupCount = 0u;
    _kittyFrameUploadBytes = 0u;
    _kittyFrameUploadCount = 0u;
    const auto keyedDrawStartedAt = std::chrono::steady_clock::now();
    drawKittyLayer(GHOSTTY_KITTY_PLACEMENT_LAYER_BELOW_BG);
    drawKittyLayer(GHOSTTY_KITTY_PLACEMENT_LAYER_BELOW_TEXT);
    drawKittyLayer(GHOSTTY_KITTY_PLACEMENT_LAYER_ABOVE_TEXT);
    const uint64_t keyedDrawUs = Debug::Perf::ElapsedUs(keyedDrawStartedAt);
    Debug::Perf::EmitValue(L"terminal.kitty.image_lookup_count", _kittyFrameImageLookupCount);
    Debug::Perf::EmitDurationUs(
        L"terminal.kitty.keyed_draw_lookup_us", std::max<uint64_t>(1u, keyedDrawUs), _kittyFrameImageLookupCount);
    Debug::Perf::EmitValue(L"terminal.kitty.frame_upload_bytes", _kittyFrameUploadBytes);
    Debug::Perf::EmitValue(L"terminal.kitty.frame_upload_count", _kittyFrameUploadCount);
    const size_t expectedHighCardinalityCacheBytes = (highCardinalityUniqueImageCount + 1u) * 4u;
    check(highCardinalityReady && cacheReuseCaptured && ! cacheReuseSubmitted && cacheReuseWork.images.empty() &&
          _kittyCachedBytes == expectedHighCardinalityCacheBytes &&
          _kittyCapturePinnedBytes == expectedHighCardinalityCacheBytes &&
          _kittyCaptureMissingImageKeys == 0u && _kittyCaptureDeferredImageKeys == 0u &&
          _kittyCaptureRejectedImageKeys == 0u && _kittyFrameImageLookupCount == _kittyCaptureVisiblePlacements &&
          _kittyFrameUploadBytes == 0u && _kittyFrameUploadCount == 0u);

    wil::unique_hwnd kittyRenderWindow(CreateWindowExW(WS_EX_TOOLWINDOW,
                                                        L"STATIC",
                                                        L"",
                                                        WS_POPUP,
                                                        0,
                                                        0,
                                                        320,
                                                        240,
                                                        nullptr,
                                                        nullptr,
                                                        g_hInstance,
                                                        nullptr));
    check(kittyRenderWindow != nullptr);
    if (kittyRenderWindow)
    {
        _windowHandle.store(kittyRenderWindow.get(), std::memory_order_release);
        const auto restoreRenderWindow = wil::scope_exit(
            [this, hwnd = kittyRenderWindow.get()]() noexcept
            {
                KillTimer(hwnd, kKittyUploadRetryTimer);
                discardDeviceResources();
                _windowHandle.store(nullptr, std::memory_order_release);
            });
        const HRESULT resourcesHr = ensureDeviceResources();
        const auto uploadImageIterator = _kittyImages.find(highCardinalityImageKey);
        check(SUCCEEDED(resourcesHr) && uploadImageIterator != _kittyImages.end() &&
              uploadImageIterator->second.state == KittyImageState::Ready);
        if (SUCCEEDED(resourcesHr) && uploadImageIterator != _kittyImages.end())
        {
            KittyImageSnapshot& uploadImage = uploadImageIterator->second;
            const auto resetUploadImage = [this](KittyImageSnapshot& image) noexcept
            {
                image.bitmap.reset();
                image.uploadState = KittyBitmapUploadState::Eligible;
                image.uploadHr = S_OK;
                image.uploadRetryAtTick = 0u;
                image.uploadDeviceGeneration = _kittyDeviceGeneration;
                image.uploadAttemptCount = 0u;
                _kittyFrameUploadBytes = 0u;
                _kittyFrameUploadCount = 0u;
                _kittyUploadDeferred = false;
                _kittyRecreateTargetAfterFrame = false;
                _debugKittyUploadFailures.clear();
            };

            _kittyUploadRetryScheduleCount = 0u;
            _kittyStableUploadFailureCount = 0u;
            _kittyDeviceRecoveryCount = 0u;
            _kittyRepaintRequestCount = 0u;
            _kittyResizeCount = 0u;
            _kittyResizeRecreateCount = 0u;
            _kittyResizeRetainedBitmapCount = 0u;
            _kittyResizeRetainedBitmapBytes = 0u;

            resetUploadImage(uploadImage);
            _debugKittyUploadFailures.push_back(HRESULT_FROM_WIN32(ERROR_RETRY));
            const KittyBitmapUploadResult oneShotFirstUpload = ensureKittyBitmap(uploadImage);
            const KittyBitmapUploadState oneShotState = uploadImage.uploadState;
            const uint32_t oneShotScheduleCount = _kittyUploadRetryScheduleCount;
            uploadImage.uploadRetryAtTick = 0u;
            const MessageRouteResult retryTimer =
                routeMouseMessage(kittyRenderWindow.get(), WM_TIMER, kKittyUploadRetryTimer, 0);
            _kittyFrameUploadBytes = 0u;
            _kittyFrameUploadCount = 0u;
            const KittyBitmapUploadResult oneShotSecondUpload = ensureKittyBitmap(uploadImage);
            check(oneShotFirstUpload.disposition == KittyBitmapUploadDisposition::RetryScheduled &&
                  oneShotFirstUpload.hr == HRESULT_FROM_WIN32(ERROR_RETRY) &&
                  oneShotState == KittyBitmapUploadState::RetryPending &&
                  oneShotScheduleCount == 1u && retryTimer.handled && _kittyRepaintRequestCount == 1u &&
                  oneShotSecondUpload.disposition == KittyBitmapUploadDisposition::Uploaded && uploadImage.bitmap &&
                  uploadImage.uploadAttemptCount == 2u);

            const auto laterImageKeyIterator = std::ranges::find_if(
                legacyImageKeys,
                [&highCardinalityImageKey](const KittyImageKey& imageKey) noexcept
                { return imageKey != highCardinalityImageKey; });
            const auto laterUploadIterator = laterImageKeyIterator == legacyImageKeys.end()
                ? _kittyImages.end()
                : _kittyImages.find(*laterImageKeyIterator);
            bool laterRetryRearmed = false;
            if (laterUploadIterator != _kittyImages.end())
            {
                KittyImageSnapshot& laterUploadImage = laterUploadIterator->second;
                resetUploadImage(uploadImage);
                resetUploadImage(laterUploadImage);
                _kittyUploadRetryScheduleCount = 0u;
                _kittyRepaintRequestCount = 0u;
                _kittyNextUploadRetryAtTick = 0u;
                _debugKittyUploadFailures.push_back(HRESULT_FROM_WIN32(ERROR_RETRY));
                _debugKittyUploadFailures.push_back(HRESULT_FROM_WIN32(ERROR_RETRY));
                const KittyBitmapUploadResult earlierFailure = ensureKittyBitmap(uploadImage);
                _kittyFrameUploadBytes = 0u;
                _kittyFrameUploadCount = 0u;
                const KittyBitmapUploadResult laterFailure = ensureKittyBitmap(laterUploadImage);
                uploadImage.uploadRetryAtTick = 0u;
                laterUploadImage.uploadRetryAtTick = GetTickCount64() + 1000u;
                const ULONGLONG laterDeadline = laterUploadImage.uploadRetryAtTick;
                const MessageRouteResult earlierTimer =
                    routeMouseMessage(kittyRenderWindow.get(), WM_TIMER, kKittyUploadRetryTimer, 0);
                _kittyFrameUploadBytes = 0u;
                _kittyFrameUploadCount = 0u;
                const KittyBitmapUploadResult earlierRetry = ensureKittyBitmap(uploadImage);
                _kittyFrameUploadBytes = 0u;
                _kittyFrameUploadCount = 0u;
                const KittyBitmapUploadResult laterBackoff = ensureKittyBitmap(laterUploadImage);
                laterRetryRearmed =
                    earlierFailure.disposition == KittyBitmapUploadDisposition::RetryScheduled &&
                    laterFailure.disposition == KittyBitmapUploadDisposition::RetryScheduled && earlierTimer.handled &&
                    earlierRetry.disposition == KittyBitmapUploadDisposition::Uploaded &&
                    laterBackoff.disposition == KittyBitmapUploadDisposition::DeferredBackoff &&
                    _kittyUploadRetryScheduleCount == 2u && _kittyNextUploadRetryAtTick == laterDeadline &&
                    laterUploadImage.uploadAttemptCount == 1u;
                KillTimer(kittyRenderWindow.get(), kKittyUploadRetryTimer);
                _kittyNextUploadRetryAtTick = 0u;
                laterUploadImage.uploadState = KittyBitmapUploadState::Eligible;
                laterUploadImage.uploadHr = S_OK;
                laterUploadImage.uploadRetryAtTick = 0u;
                _debugKittyUploadFailures.clear();
            }
            check(laterRetryRearmed);
            _kittyUploadRetryScheduleCount = oneShotScheduleCount;
            _kittyRepaintRequestCount = 1u;

            resetUploadImage(uploadImage);
            _debugKittyUploadFailures.push_back(E_OUTOFMEMORY);
            _debugKittyUploadFailures.push_back(E_OUTOFMEMORY);
            const KittyBitmapUploadResult firstOomUpload = ensureKittyBitmap(uploadImage);
            _kittyFrameUploadBytes = 0u;
            _kittyFrameUploadCount = 0u;
            const KittyBitmapUploadResult secondOomUpload = ensureKittyBitmap(uploadImage);
            check(firstOomUpload.disposition == KittyBitmapUploadDisposition::PermanentFailure &&
                  secondOomUpload.disposition == KittyBitmapUploadDisposition::PermanentFailure &&
                  uploadImage.uploadState == KittyBitmapUploadState::PermanentFailure &&
                  uploadImage.uploadHr == E_OUTOFMEMORY && uploadImage.uploadAttemptCount == 1u &&
                  _debugKittyUploadFailures.size() == 1u && _kittyStableUploadFailureCount == 1u &&
                  _kittyRepaintRequestCount == 1u);

            resetUploadImage(uploadImage);
            _debugKittyUploadFailures.push_back(D2DERR_RECREATE_TARGET);
            const KittyBitmapUploadResult recreateUpload = ensureKittyBitmap(uploadImage);
            const bool recoveredAfterFrame = recoverKittyDeviceAfterFrame(kittyRenderWindow.get());
            check(recreateUpload.disposition == KittyBitmapUploadDisposition::RecreateTarget &&
                  uploadImage.uploadState == KittyBitmapUploadState::RecreatePending &&
                  uploadImage.uploadHr == D2DERR_RECREATE_TARGET && uploadImage.uploadAttemptCount == 1u &&
                  recoveredAfterFrame && ! _renderTarget && _kittyDeviceRecoveryCount == 1u &&
                  _kittyRepaintRequestCount == 2u);

            const HRESULT recreatedResourcesHr = ensureDeviceResources();
            resetUploadImage(uploadImage);
            const KittyBitmapUploadResult initialResizeUpload =
                SUCCEEDED(recreatedResourcesHr) ? ensureKittyBitmap(uploadImage) : KittyBitmapUploadResult{};
            ID2D1Bitmap* const retainedBitmap = uploadImage.bitmap.get();
            bool resizeStormPreserved =
                initialResizeUpload.disposition == KittyBitmapUploadDisposition::Uploaded && retainedBitmap != nullptr;
            constexpr uint32_t resizeSamples = 32u;
            for (uint32_t sample = 0u; sample < resizeSamples; ++sample)
            {
                const int width = 321 + static_cast<int>(sample % 7u);
                const int height = 241 + static_cast<int>(sample % 5u);
                const bool moved = SetWindowPos(kittyRenderWindow.get(),
                                                nullptr,
                                                0,
                                                0,
                                                width,
                                                height,
                                                SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOOWNERZORDER | SWP_NOZORDER) != FALSE;
                const MessageRouteResult resizeResult = routeRenderMessage(
                    kittyRenderWindow.get(), WM_SIZE, SIZE_RESTORED, MAKELPARAM(width, height));
                resizeStormPreserved = resizeStormPreserved && moved && resizeResult.handled && _renderTarget &&
                    uploadImage.bitmap.get() == retainedBitmap;
            }
            check(resizeStormPreserved && _kittyResizeCount == resizeSamples &&
                  _kittyResizeRecreateCount == 0u && _kittyResizeRetainedBitmapCount >= 1u &&
                  _kittyResizeRetainedBitmapBytes >= uploadImage.bgra.size() && uploadImage.uploadAttemptCount == 1u);

            Debug::Perf::EmitValue(L"terminal.kitty.upload_retry_scheduled_count", _kittyUploadRetryScheduleCount);
            Debug::Perf::EmitValue(L"terminal.kitty.upload_stable_failure_count", _kittyStableUploadFailureCount);
            Debug::Perf::EmitValue(L"terminal.kitty.device_recovery_count", _kittyDeviceRecoveryCount);
            Debug::Perf::EmitValue(L"terminal.kitty.repaint_request_count", _kittyRepaintRequestCount);
            Debug::Perf::EmitValue(L"terminal.kitty.resize_count", _kittyResizeCount);
            Debug::Perf::EmitValue(L"terminal.kitty.resize_recreate_count", _kittyResizeRecreateCount);
            Debug::Perf::EmitValue(L"terminal.kitty.resize_retained_bitmap_count", _kittyResizeRetainedBitmapCount);
            Debug::Perf::EmitValue(L"terminal.kitty.resize_retained_bitmap_bytes", _kittyResizeRetainedBitmapBytes);
        }
    }
    return *failedTests == 0u ? S_OK : E_FAIL;
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderTerminalDebugSelfTests(
    unsigned int* passedTests,
    unsigned int* failedTests) noexcept
{
    if (passedTests == nullptr || failedTests == nullptr)
    {
        return E_POINTER;
    }
    *passedTests = 0u;
    *failedTests = 0u;
    const auto check = [passedTests, failedTests](bool condition) noexcept
    {
        if (condition)
        {
            ++*passedTests;
        }
        else
        {
            ++*failedTests;
        }
    };

    check(GhosttyKeyFromVirtualKey(VK_RETURN) == GHOSTTY_KEY_ENTER);

    bool translatedCharacterSuppression = false;
    std::string selectedCtrlCBytes;
    MarkTranslatedCharacterConsumed(translatedCharacterSuppression);
    if (! ConsumeTranslatedCharacterSuppression(translatedCharacterSuppression))
    {
        selectedCtrlCBytes.push_back('\x03');
    }
    check(selectedCtrlCBytes.empty() && ! ConsumeTranslatedCharacterSuppression(translatedCharacterSuppression));

    std::string unselectedCtrlCBytes;
    // The keydown-time decision is final: releasing Ctrl/Shift before the queued
    // WM_CHAR cannot suppress an unselected Ctrl+C ETX.
    if (! ConsumeTranslatedCharacterSuppression(translatedCharacterSuppression))
    {
        unselectedCtrlCBytes.push_back('\x03');
    }
    check(unselectedCtrlCBytes == "\x03");
    check(VirtualKeyProducesTranslatedCharacter(VK_RETURN));
    check(VirtualKeyProducesTranslatedCharacter(VK_ESCAPE));
    check(VirtualKeyProducesTranslatedCharacter(L'7'));
    check(VirtualKeyProducesTranslatedCharacter(VK_OEM_PLUS));
    check(VirtualKeyProducesTranslatedCharacter(VK_NUMPAD3));
    check(! VirtualKeyProducesTranslatedCharacter(VK_UP));
    check(! VirtualKeyProducesTranslatedCharacter(VK_HOME));
    check(! VirtualKeyProducesTranslatedCharacter(VK_END));
    check(! VirtualKeyProducesTranslatedCharacter(VK_F5));

#if defined(ENABLE_TESTS)
    {
        Terminal routingTerminal;
        unsigned int routingPassed = 0u;
        unsigned int routingFailed = 0u;
        const HRESULT routingHr =
            routingTerminal.DebugRunInputRoutingContractSelfTests(&routingPassed, &routingFailed);
        *passedTests += routingPassed;
        *failedTests += routingFailed;
        check(SUCCEEDED(routingHr));
    }
    {
        Terminal runtimeLifecycleTerminal;
        unsigned int lifecyclePassed = 0u;
        unsigned int lifecycleFailed = 0u;
        const HRESULT lifecycleHr = runtimeLifecycleTerminal.DebugRunRuntimeLifecycleSelfTests(
            &lifecyclePassed, &lifecycleFailed);
        *passedTests += lifecyclePassed;
        *failedTests += lifecycleFailed;
        check(SUCCEEDED(lifecycleHr));
    }
#endif

    const std::wstring extendedDrivePath = LR"(\\?\C:\terminal)";
    TerminalLogicalLocation extendedDrive{};
    extendedDrive.sizeBytes = sizeof(extendedDrive);
    extendedDrive.kind = TerminalLocationKind::WindowsLocal;
    extendedDrive.windowsPath = {extendedDrivePath.data(), static_cast<uint32_t>(extendedDrivePath.size())};
    check(ValidateTerminalLogicalLocation(extendedDrive));

    const std::wstring extendedUncPath = LR"(\\?\UNC\server\share\terminal)";
    TerminalLogicalLocation extendedUnc{};
    extendedUnc.sizeBytes = sizeof(extendedUnc);
    extendedUnc.kind = TerminalLocationKind::WindowsUnc;
    extendedUnc.windowsPath = {extendedUncPath.data(), static_cast<uint32_t>(extendedUncPath.size())};
    check(ValidateTerminalLogicalLocation(extendedUnc));

    const std::wstring devicePath = LR"(\\.\PhysicalDrive0)";
    TerminalLogicalLocation device{};
    device.sizeBytes = sizeof(device);
    device.kind = TerminalLocationKind::WindowsUnc;
    device.windowsPath = {devicePath.data(), static_cast<uint32_t>(devicePath.size())};
    check(! ValidateTerminalLogicalLocation(device));

    const std::wstring insertionParent = L"C:\\work";
    const std::wstring insertionItem = L"C:\\work\\item name.txt";
    const std::wstring insertionLeaf = L"item name.txt";
    const std::wstring insertionCurrent = L"C:\\current";
    TerminalPathInsertion insertion{};
    insertion.mode = TerminalPathInsertionMode::ContextualLeafOrFull;
    insertion.parentLocation.windowsPath = {insertionParent.data(), static_cast<uint32_t>(insertionParent.size())};
    insertion.itemLocation.windowsPath = {insertionItem.data(), static_cast<uint32_t>(insertionItem.size())};
    insertion.displayLeaf = {insertionLeaf.data(), static_cast<uint32_t>(insertionLeaf.size())};
    check(ResolvePathInsertionValue(insertion, true, true, false, insertionParent) == insertionLeaf);
    check(ResolvePathInsertionValue(insertion, true, true, false, L"C:\\other") == insertionItem);
    check(ResolvePathInsertionValue(insertion, true, true, true, insertionParent) == insertionItem);
    insertion.mode = TerminalPathInsertionMode::AlwaysFull;
    check(ResolvePathInsertionValue(insertion, true, true, true, insertionParent) == insertionItem);
    insertion.mode = TerminalPathInsertionMode::CurrentDirectoryFull;
    insertion.initiatingSourceLocation.windowsPath = {
        insertionCurrent.data(), static_cast<uint32_t>(insertionCurrent.size())};
    check(ResolvePathInsertionValue(insertion, true, true, true, insertionParent) == insertionCurrent);

    const TerminalCommandSurfaceModel::FindResult findResult = TerminalCommandSurfaceModel::Find(
        L"alpha\r\nBravo ALPHA\r\ncharlie\r\nalpha delta", L"alpha", 16u);
    check(findResult.scannedLines == 4u && findResult.totalMatches == 3u && findResult.rows.size() == 3u &&
          findResult.rows[1].lineNumber == 1u && findResult.rows[1].matchOffset == 6u);
    std::stop_source cancelledFind;
    cancelledFind.request_stop();
    check(TerminalCommandSurfaceModel::Find(L"alpha", L"alpha", 1u, cancelledFind.get_token()).cancelled);

    const std::vector<std::wstring> persistedHistory{L"git status", L"cmake --build .", L"git log"};
    const std::vector<std::wstring> currentHistory{L"git status", L"git switch feature", L"git diff"};
    const TerminalCommandSurfaceModel::SuggestionResult suggestions = TerminalCommandSurfaceModel::FilterSuggestions(
        persistedHistory, currentHistory, L"git", 16u);
    check(suggestions.rows.size() == 4u && suggestions.rows[0].currentSession && suggestions.rows[0].text == L"git diff" &&
          suggestions.rows[1].text == L"git switch feature" && suggestions.rows[2].text == L"git status" &&
          suggestions.rows[3].text == L"git log");
    const std::vector<std::wstring> unicodePersisted{L"ÉCHO", L"ΔΟΚΙΜΗ", L"ПРИВЕТ"};
    const std::vector<std::wstring> unicodeCurrent{L"écho", L"δοκιμη", L"привет"};
    const TerminalCommandSurfaceModel::SuggestionResult unicodeSuggestions =
        TerminalCommandSurfaceModel::FilterSuggestions(unicodePersisted, unicodeCurrent, L"", 16u);
    check(unicodeSuggestions.rows.size() == 3u &&
          std::ranges::any_of(unicodeSuggestions.rows, [](const auto& row) noexcept { return row.text == L"écho"; }) &&
          std::ranges::any_of(unicodeSuggestions.rows, [](const auto& row) noexcept { return row.text == L"δοκιμη"; }) &&
          std::ranges::any_of(unicodeSuggestions.rows, [](const auto& row) noexcept { return row.text == L"привет"; }));
    const wchar_t* currentLocale = _wsetlocale(LC_CTYPE, nullptr);
    const std::wstring originalLocale = currentLocale ? currentLocale : L"";
    const auto restoreLocale = wil::scope_exit([&]() noexcept
    {
        if (! originalLocale.empty())
        {
            static_cast<void>(_wsetlocale(LC_CTYPE, originalLocale.c_str()));
        }
    });
    const std::array<const wchar_t*, 3u> turkishLocaleNames{
        L"Turkish_Turkey.1254", L"tr-TR", L"tr_TR"};
    bool appliedTurkishLocale = false;
    for (const wchar_t* localeName : turkishLocaleNames)
    {
        if (_wsetlocale(LC_CTYPE, localeName) != nullptr)
        {
            appliedTurkishLocale = true;
            break;
        }
    }
    if (appliedTurkishLocale)
    {
        const TerminalCommandSurfaceModel::SuggestionResult localizedSuggestions =
            TerminalCommandSurfaceModel::FilterSuggestions(unicodePersisted, unicodeCurrent, L"", 16u);
        check(localizedSuggestions.rows.size() == unicodeSuggestions.rows.size() &&
              std::ranges::equal(localizedSuggestions.rows,
                                 unicodeSuggestions.rows,
                                 [](const auto& left, const auto& right) noexcept
                                 {
                                     return left.text == right.text && left.currentSession == right.currentSession;
                                 }));
    }
    else
    {
        Debug::Warning(L"TerminalSelfTest: Turkish C locale unavailable; ordinal Unicode fixtures still executed.");
    }
    check(TerminalCommandSurfaceModel::IsSafeSuggestion(L"Get-ChildItem") &&
          !TerminalCommandSurfaceModel::IsSafeSuggestion(L"hidden\r\ncommand"));

    const std::wstring packedSessionHistory(L"git status\0.\\build.ps1", 22u);
    std::array<std::string, 4u> integrationFields;
    const bool integrationFieldsEncoded =
        SUCCEEDED(EncodeIntegrationField(L"C:\\work", integrationFields[0])) &&
        SUCCEEDED(EncodeIntegrationField(L"C:\\profile\\ConsoleHost_history.txt", integrationFields[1])) &&
        SUCCEEDED(EncodeIntegrationField(L"buil", integrationFields[2])) &&
        SUCCEEDED(EncodeIntegrationField(packedSessionHistory, integrationFields[3]));
    IntegrationRequest integrationRequest;
    const std::string integrationRecord = std::format("RS2\t{}\t{}\t{}\t{}",
                                                      integrationFields[0],
                                                      integrationFields[1],
                                                      integrationFields[2],
                                                      integrationFields[3]);
    check(integrationFieldsEncoded && SUCCEEDED(ParseIntegrationRequest(integrationRecord, integrationRequest)) &&
          integrationRequest.versioned && integrationRequest.currentDirectory == L"C:\\work" &&
          integrationRequest.historyPath == L"C:\\profile\\ConsoleHost_history.txt" &&
          integrationRequest.promptBuffer == L"buil" && integrationRequest.currentSessionHistory.size() == 2u &&
          integrationRequest.currentSessionHistory[0] == L"git status" &&
          integrationRequest.currentSessionHistory[1] == L".\\build.ps1");
    IntegrationRequest invalidIntegrationRequest;
    check(ParseIntegrationRequest("RS2\t%%%\t\t\t", invalidIntegrationRequest) == E_INVALIDARG);

    const SCROLLINFO nativeScrollInfo = MakeTerminalScrollInfo(
        GhosttyTerminalScrollbar{.total = 100u, .offset = 25u, .len = 10u});
    check(nativeScrollInfo.nMin == 0 && nativeScrollInfo.nMax == 99 && nativeScrollInfo.nPage == 10u &&
          nativeScrollInfo.nPos == 25 && (nativeScrollInfo.fMask & SIF_DISABLENOSCROLL) != 0u);

    const TerminalWheelScrollResult doubleDetent = ResolveTerminalWheelScroll(0, WHEEL_DELTA * 2, 3u, 24u);
    const TerminalWheelScrollResult partialDetent = ResolveTerminalWheelScroll(0, WHEEL_DELTA / 2, 3u, 24u);
    const TerminalWheelScrollResult completedDetent = ResolveTerminalWheelScroll(
        partialDetent.remainder, WHEEL_DELTA / 2, 3u, 24u);
    const TerminalWheelScrollResult pageDetents = ResolveTerminalWheelScroll(0, -WHEEL_DELTA * 2, WHEEL_PAGESCROLL, 24u);
    check(doubleDetent.rows == -6 && doubleDetent.remainder == 0 &&
          partialDetent.rows == 0 && partialDetent.remainder == WHEEL_DELTA / 2 &&
          completedDetent.rows == -3 && completedDetent.remainder == 0 &&
          pageDetents.rows == 48 && pageDetents.remainder == 0);

    const TerminalPngDecodeThreadContext pngDecodeContext;
    check(SUCCEEDED(pngDecodeContext.Status()));
    TerminalRuntimeLoader runtime;
    const HRESULT runtimeLoadHr = runtime.Load();
    if (FAILED(runtimeLoadHr))
    {
        check(false);
        return runtimeLoadHr;
    }

    GhosttyTerminal terminal = nullptr;
    GhosttyRenderState renderState = nullptr;
    GhosttyRenderStateRowIterator rows = nullptr;
    GhosttyRenderStateRowCells cells = nullptr;
    GhosttyKeyEncoder encoder = nullptr;
    GhosttyKeyEvent event = nullptr;
    GhosttyMouseEncoder mouseEncoder = nullptr;
    GhosttyMouseEvent mouseEvent = nullptr;
    GhosttyKittyGraphicsPlacementIterator kittyPlacements = nullptr;
    GhosttyTerminal placementBoundTerminal = nullptr;
    GhosttyTrackedGridRef trackedCell = nullptr;
    const auto cleanup = wil::scope_exit([&] noexcept
    {
        if (mouseEvent != nullptr)
        {
            runtime.mouseEventFree(mouseEvent);
        }
        if (kittyPlacements != nullptr)
        {
            runtime.kittyPlacementIteratorFree(kittyPlacements);
        }
        if (trackedCell != nullptr)
        {
            runtime.trackedGridRefFree(trackedCell);
        }
        if (mouseEncoder != nullptr)
        {
            runtime.mouseEncoderFree(mouseEncoder);
        }
        if (event != nullptr)
        {
            runtime.keyEventFree(event);
        }
        if (encoder != nullptr)
        {
            runtime.keyEncoderFree(encoder);
        }
        if (cells != nullptr)
        {
            runtime.renderRowCellsFree(cells);
        }
        if (rows != nullptr)
        {
            runtime.renderRowIteratorFree(rows);
        }
        if (renderState != nullptr)
        {
            runtime.renderStateFree(renderState);
        }
        if (terminal != nullptr)
        {
            runtime.terminalFree(terminal);
        }
        if (placementBoundTerminal != nullptr)
        {
            runtime.terminalFree(placementBoundTerminal);
        }
        runtime.Unload();
    });

    const bool created = runtime.terminalNew(nullptr, &terminal, 8u, 2u) == GHOSTTY_SUCCESS && terminal != nullptr &&
        runtime.renderStateNew(nullptr, &renderState) == GHOSTTY_SUCCESS && renderState != nullptr &&
        runtime.renderRowIteratorNew(nullptr, &rows) == GHOSTTY_SUCCESS && rows != nullptr &&
        runtime.renderRowCellsNew(nullptr, &cells) == GHOSTTY_SUCCESS && cells != nullptr &&
        runtime.keyEncoderNew(nullptr, &encoder) == GHOSTTY_SUCCESS && encoder != nullptr &&
        runtime.keyEventNew(nullptr, &event) == GHOSTTY_SUCCESS && event != nullptr &&
        runtime.mouseEncoderNew(nullptr, &mouseEncoder) == GHOSTTY_SUCCESS && mouseEncoder != nullptr &&
        runtime.mouseEventNew(nullptr, &mouseEvent) == GHOSTTY_SUCCESS && mouseEvent != nullptr &&
        runtime.kittyPlacementIteratorNew(nullptr, &kittyPlacements) == GHOSTTY_SUCCESS && kittyPlacements != nullptr;
    check(created);
    if (! created)
    {
        return E_FAIL;
    }

    const auto writeVt = [&](std::string_view bytes) noexcept
    {
        runtime.terminalVtWrite(
            terminal, reinterpret_cast<const uint8_t*>(bytes.data()), bytes.size());
    };
    const auto modeEquals = [&](GhosttyMode mode, bool expected) noexcept
    {
        bool value = ! expected;
        return runtime.GetTerminalMode(terminal, mode, value) == GHOSTTY_SUCCESS && value == expected;
    };

    check(modeEquals(GHOSTTY_MODE_DECCKM, false) &&
          modeEquals(GHOSTTY_MODE_NORMAL_MOUSE, false) &&
          modeEquals(GHOSTTY_MODE_SGR_MOUSE, false) &&
          modeEquals(GHOSTTY_MODE_FOCUS_EVENT, false) &&
          modeEquals(GHOSTTY_MODE_ALT_SCREEN_SAVE, false) &&
          modeEquals(GHOSTTY_MODE_BRACKETED_PASTE, false));
    constexpr std::string_view enableModes =
        "\x1b[?1h\x1b[?1000h\x1b[?1006h\x1b[?1004h\x1b[?1049h\x1b[?2004h";
    writeVt(enableModes);
    check(modeEquals(GHOSTTY_MODE_DECCKM, true) &&
          modeEquals(GHOSTTY_MODE_NORMAL_MOUSE, true) &&
          modeEquals(GHOSTTY_MODE_SGR_MOUSE, true) &&
          modeEquals(GHOSTTY_MODE_FOCUS_EVENT, true) &&
          modeEquals(GHOSTTY_MODE_ALT_SCREEN_SAVE, true) &&
          modeEquals(GHOSTTY_MODE_BRACKETED_PASTE, true));
    constexpr std::string_view disableModes =
        "\x1b[?2004l\x1b[?1049l\x1b[?1004l\x1b[?1006l\x1b[?1000l\x1b[?1l";
    writeVt(disableModes);
    check(modeEquals(GHOSTTY_MODE_DECCKM, false) &&
          modeEquals(GHOSTTY_MODE_NORMAL_MOUSE, false) &&
          modeEquals(GHOSTTY_MODE_SGR_MOUSE, false) &&
          modeEquals(GHOSTTY_MODE_FOCUS_EVENT, false) &&
          modeEquals(GHOSTTY_MODE_ALT_SCREEN_SAVE, false) &&
          modeEquals(GHOSTTY_MODE_BRACKETED_PASTE, false));

    // A UTF-8 continuation byte in the C1 range must remain DCS payload and
    // must not escape into a live CSI that renders the payload's trailing X.
    const std::array<uint8_t, 12u> dcsHighBytes{
        0x1Bu, 'P', 'q', 0xC3u, 0x9Bu, '3', '1', 'm', 'X', 0x1Bu, '\\', 'Y'};
    runtime.terminalVtWrite(terminal, dcsHighBytes.data(), dcsHighBytes.size());
    const bool dcsSnapshotUpdated = runtime.renderStateBeginUpdate(renderState, terminal) == GHOSTTY_SUCCESS &&
        runtime.renderStateEndUpdate(renderState) == GHOSTTY_SUCCESS;
    bool dcsFirstCellReady = false;
    if (dcsSnapshotUpdated &&
        runtime.renderStateGet(renderState, GHOSTTY_RENDER_STATE_DATA_ROW_ITERATOR, &rows) == GHOSTTY_SUCCESS &&
        runtime.renderRowIteratorNext(rows) &&
        runtime.renderRowGet(rows, GHOSTTY_RENDER_STATE_ROW_DATA_CELLS, &cells) == GHOSTTY_SUCCESS &&
        runtime.renderRowCellsNext(cells))
    {
        dcsFirstCellReady = true;
    }
    std::array<uint8_t, 8u> dcsGraphemeBytes{};
    GhosttyBuffer dcsGrapheme{dcsGraphemeBytes.data(), dcsGraphemeBytes.size(), 0u};
    check(dcsFirstCellReady &&
          runtime.renderRowCellsGet(cells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_GRAPHEMES_UTF8, &dcsGrapheme) == GHOSTTY_SUCCESS &&
          std::string_view(reinterpret_cast<const char*>(dcsGrapheme.ptr), dcsGrapheme.len) == "Y");
    writeVt("\x1b" "c");

    constexpr std::string_view styledText = "\x1b[31;44;1mA\x1b[0m";
    runtime.terminalVtWrite(
        terminal, reinterpret_cast<const uint8_t*>(styledText.data()), styledText.size());
    const bool snapshotUpdated = runtime.renderStateBeginUpdate(renderState, terminal) == GHOSTTY_SUCCESS &&
        runtime.renderStateEndUpdate(renderState) == GHOSTTY_SUCCESS;
    check(snapshotUpdated);

    bool firstCellReady = false;
    if (snapshotUpdated &&
        runtime.renderStateGet(renderState, GHOSTTY_RENDER_STATE_DATA_ROW_ITERATOR, &rows) == GHOSTTY_SUCCESS &&
        runtime.renderRowIteratorNext(rows) &&
        runtime.renderRowGet(rows, GHOSTTY_RENDER_STATE_ROW_DATA_CELLS, &cells) == GHOSTTY_SUCCESS &&
        runtime.renderRowCellsNext(cells))
    {
        firstCellReady = true;
    }
    check(firstCellReady);
    if (firstCellReady)
    {
        std::array<uint8_t, 16u> graphemeBytes{};
        GhosttyBuffer grapheme{graphemeBytes.data(), graphemeBytes.size(), 0u};
        GhosttyStyle style{};
        style.size = sizeof(style);
        GhosttyColorRgb foreground{};
        GhosttyColorRgb background{};
        check(runtime.renderRowCellsGet(cells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_GRAPHEMES_UTF8, &grapheme) == GHOSTTY_SUCCESS &&
              std::string_view(reinterpret_cast<const char*>(grapheme.ptr), grapheme.len) == "A");
        check(runtime.renderRowCellsGet(cells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_STYLE, &style) == GHOSTTY_SUCCESS && style.bold);
        check(runtime.renderRowCellsGet(cells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_FG_COLOR, &foreground) == GHOSTTY_SUCCESS &&
              runtime.renderRowCellsGet(cells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_BG_COLOR, &background) == GHOSTTY_SUCCESS &&
              (foreground.r != background.r || foreground.g != background.g || foreground.b != background.b));
    }

    constexpr std::string_view hyperlinkText =
        "\r\n\x1b]8;;https://example.test/path\x1b\\L\x1b]8;;\x1b\\";
    runtime.terminalVtWrite(
        terminal, reinterpret_cast<const uint8_t*>(hyperlinkText.data()), hyperlinkText.size());
    GhosttyPoint hyperlinkPoint{};
    hyperlinkPoint.tag = GHOSTTY_POINT_TAG_VIEWPORT;
    hyperlinkPoint.value.coordinate.x = 0u;
    hyperlinkPoint.value.coordinate.y = 1u;
    GhosttyGridRef hyperlinkCell{};
    hyperlinkCell.size = sizeof(hyperlinkCell);
    std::array<uint8_t, 128u> hyperlinkUri{};
    size_t hyperlinkLength = 0u;
    check(runtime.terminalGridRef(terminal, hyperlinkPoint, &hyperlinkCell) == GHOSTTY_SUCCESS &&
          runtime.gridRefHyperlinkUri(
              &hyperlinkCell, hyperlinkUri.data(), hyperlinkUri.size(), &hyperlinkLength) == GHOSTTY_SUCCESS &&
          std::string_view(reinterpret_cast<const char*>(hyperlinkUri.data()), hyperlinkLength) ==
              "https://example.test/path");

    std::vector<uint8_t> clipboardBoundaryBytes(kGhosttyClipboardWriteMaxBytes + 1u, static_cast<uint8_t>('x'));
    constexpr std::string_view clipboardMime = "text/plain";
    GhosttyClipboardContent boundaryContent{};
    boundaryContent.mime.ptr = reinterpret_cast<const uint8_t*>(clipboardMime.data());
    boundaryContent.mime.len = clipboardMime.size();
    boundaryContent.data.ptr = clipboardBoundaryBytes.data();
    GhosttyClipboardWrite boundaryWrite{};
    boundaryWrite.size = sizeof(boundaryWrite);
    boundaryWrite.location = GHOSTTY_CLIPBOARD_LOCATION_STANDARD;
    boundaryWrite.contents = &boundaryContent;
    boundaryWrite.contents_len = 1u;
    GhosttyClipboardDataView boundaryView{};
    boundaryWrite.size = offsetof(GhosttyClipboardWrite, contents_len) + sizeof(size_t) - 1u;
    check(ValidateGhosttyClipboardWrite(
              &boundaryWrite, kGhosttyClipboardWriteMaxBytes, boundaryView) == GHOSTTY_CLIPBOARD_WRITE_RESULT_UNSUPPORTED);
    boundaryWrite.size = sizeof(boundaryWrite);
    boundaryContent.data.ptr = nullptr;
    boundaryContent.data.len = 1u;
    check(ValidateGhosttyClipboardWrite(
              &boundaryWrite, kGhosttyClipboardWriteMaxBytes, boundaryView) == GHOSTTY_CLIPBOARD_WRITE_RESULT_INVALID_DATA);
    boundaryContent.data.ptr = clipboardBoundaryBytes.data();
    for (const size_t length : {
             kGhosttyClipboardWriteMaxBytes - 1u,
             kGhosttyClipboardWriteMaxBytes,
             kGhosttyClipboardWriteMaxBytes + 1u})
    {
        boundaryContent.data.len = length;
        const GhosttyClipboardWriteResult expected = length <= kGhosttyClipboardWriteMaxBytes
            ? GHOSTTY_CLIPBOARD_WRITE_RESULT_SUCCESS
            : GHOSTTY_CLIPBOARD_WRITE_RESULT_INVALID_DATA;
        check(ValidateGhosttyClipboardWrite(
                  &boundaryWrite, kGhosttyClipboardWriteMaxBytes, boundaryView) == expected &&
              (expected != GHOSTTY_CLIPBOARD_WRITE_RESULT_SUCCESS || boundaryView.length == length));
    }
    constexpr std::array<uint8_t, 2u> invalidClipboardUtf8{{0xC3u, 0x28u}};
    check(! Common::Strings::TryUtf16FromUtf8Strict(std::string_view(
               reinterpret_cast<const char*>(invalidClipboardUtf8.data()), invalidClipboardUtf8.size())).has_value());

    RuntimeClipboardCapture clipboardCapture{};
    constexpr std::string_view clipboardWrite = "\x1b]52;c;Y2xpcGJvYXJkLXRlc3Q=\x1b\\";
    bool clipboardCallbackConfigured =
        runtime.terminalSet(terminal, GHOSTTY_TERMINAL_OPT_USERDATA, &clipboardCapture) == GHOSTTY_SUCCESS;
    const size_t clipboardWriteMaxBytes = kGhosttyClipboardWriteMaxBytes;
    const bool kittyClipboardWriteEnabled = false;
    clipboardCallbackConfigured = clipboardCallbackConfigured &&
        runtime.terminalSet(terminal, GHOSTTY_TERMINAL_OPT_WRITE_PTY,
                            reinterpret_cast<const void*>(&CaptureRuntimePty)) == GHOSTTY_SUCCESS &&
        runtime.terminalSet(terminal, GHOSTTY_TERMINAL_OPT_CLIPBOARD_READ, nullptr) == GHOSTTY_SUCCESS &&
        runtime.terminalSet(terminal, GHOSTTY_TERMINAL_OPT_CLIPBOARD_WRITE_MAX_BYTES, &clipboardWriteMaxBytes) == GHOSTTY_SUCCESS &&
        runtime.terminalSet(terminal, GHOSTTY_TERMINAL_OPT_KITTY_CLIPBOARD_WRITE_ENABLED, &kittyClipboardWriteEnabled) == GHOSTTY_SUCCESS;
    check(! GhosttySizedFieldPresent(
        offsetof(GhosttyClipboardWrite, reply), offsetof(GhosttyClipboardWrite, reply), sizeof(boundaryWrite.reply)));
    clipboardCallbackConfigured = clipboardCallbackConfigured &&
        runtime.terminalSet(terminal,
                            GHOSTTY_TERMINAL_OPT_CLIPBOARD_WRITE,
                            reinterpret_cast<const void*>(&CaptureRuntimeClipboard)) == GHOSTTY_SUCCESS;
    runtime.terminalVtWrite(
        terminal, reinterpret_cast<const uint8_t*>(clipboardWrite.data()), clipboardWrite.size());
    check(clipboardCallbackConfigured && clipboardCapture.invoked && clipboardCapture.callbackCount == 1u &&
          clipboardCapture.location == GHOSTTY_CLIPBOARD_LOCATION_STANDARD &&
          std::string_view(reinterpret_cast<const char*>(clipboardCapture.data.data()), clipboardCapture.length) ==
              "clipboard-test");
    check(clipboardCapture.replyCount == 1u &&
          clipboardCapture.replyResult == GHOSTTY_CLIPBOARD_WRITE_RESULT_SUCCESS);
    clipboardCapture = {};
    constexpr std::string_view osc1337Write = "\x1b]1337;Copy=:aVRlcm0y\x1b\\";
    runtime.terminalVtWrite(
        terminal, reinterpret_cast<const uint8_t*>(osc1337Write.data()), osc1337Write.size());
    check(clipboardCapture.invoked && clipboardCapture.callbackCount == 1u && clipboardCapture.replyCount == 1u &&
          std::string_view(reinterpret_cast<const char*>(clipboardCapture.data.data()), clipboardCapture.length) == "iTerm2");

    clipboardCapture = {};
    constexpr std::string_view kittyWriteBegin = "\x1b]5522;type=write:id=x\x1b\\";
    runtime.terminalVtWrite(
        terminal, reinterpret_cast<const uint8_t*>(kittyWriteBegin.data()), kittyWriteBegin.size());
    constexpr std::string_view kittyWriteDenied = "\x1b]5522;type=write:status=ENOSYS:id=x\x1b\\";
    check(clipboardCapture.callbackCount == 0u &&
          std::string_view(reinterpret_cast<const char*>(clipboardCapture.ptyData.data()), clipboardCapture.ptyLength) ==
              kittyWriteDenied);
    constexpr std::string_view kittyWriteRemainder =
        "\x1b]5522;type=wdata:mime=dGV4dC9wbGFpbg==;R2hvc3Q=\x1b\\"
        "\x1b]5522;type=wdata\x1b\\";
    runtime.terminalVtWrite(
        terminal, reinterpret_cast<const uint8_t*>(kittyWriteRemainder.data()), kittyWriteRemainder.size());
    check(clipboardCapture.callbackCount == 0u &&
          std::string_view(reinterpret_cast<const char*>(clipboardCapture.ptyData.data()), clipboardCapture.ptyLength) ==
              kittyWriteDenied);

    clipboardCapture = {};
    constexpr std::string_view osc52Read = "\x1b]52;c;?\x1b\\";
    runtime.terminalVtWrite(terminal, reinterpret_cast<const uint8_t*>(osc52Read.data()), osc52Read.size());
    constexpr std::string_view kittyRead = "\x1b]5522;type=read:id=r;dGV4dC9wbGFpbg==\x1b\\";
    runtime.terminalVtWrite(terminal, reinterpret_cast<const uint8_t*>(kittyRead.data()), kittyRead.size());
    constexpr std::string_view kittyReadDenied = "\x1b]5522;type=read:status=EPERM:id=r\x1b\\";
    check(clipboardCapture.callbackCount == 0u &&
          std::string_view(reinterpret_cast<const char*>(clipboardCapture.ptyData.data()), clipboardCapture.ptyLength) ==
              kittyReadDenied);
    static_cast<void>(runtime.terminalSet(terminal, GHOSTTY_TERMINAL_OPT_WRITE_PTY, nullptr));
    static_cast<void>(runtime.terminalSet(terminal, GHOSTTY_TERMINAL_OPT_CLIPBOARD_WRITE, nullptr));
    static_cast<void>(runtime.terminalSet(terminal, GHOSTTY_TERMINAL_OPT_USERDATA, nullptr));

    runtime.keyEventSetAction(event, GHOSTTY_KEY_ACTION_PRESS);
    runtime.keyEventSetKey(event, GHOSTTY_KEY_ARROW_UP);
    runtime.keyEventSetMods(event, 0u);
    runtime.keyEventSetUtf8(event, nullptr, 0u);
    runtime.keyEventSetUnshiftedCodepoint(event, 0u);
    runtime.keyEncoderSetFromTerminal(encoder, terminal);
    std::array<char, 32u> normalBytes{};
    size_t normalLength = 0u;
    const GhosttyResult normalResult = runtime.keyEncoderEncode(
        encoder, event, normalBytes.data(), normalBytes.size(), &normalLength);

    constexpr std::string_view applicationCursor = "\x1b[?1h";
    runtime.terminalVtWrite(
        terminal, reinterpret_cast<const uint8_t*>(applicationCursor.data()), applicationCursor.size());
    runtime.keyEncoderSetFromTerminal(encoder, terminal);
    std::array<char, 32u> applicationBytes{};
    size_t applicationLength = 0u;
    const GhosttyResult applicationResult = runtime.keyEncoderEncode(
        encoder, event, applicationBytes.data(), applicationBytes.size(), &applicationLength);
    check(normalResult == GHOSTTY_SUCCESS && applicationResult == GHOSTTY_SUCCESS && normalLength != 0u && applicationLength != 0u &&
          std::string_view(normalBytes.data(), normalLength) != std::string_view(applicationBytes.data(), applicationLength));

    std::array<char, 32u> pasteSource{'a', '\n', 'b'};
    std::array<char, 64u> pasteOutput{};
    size_t pasteLength = 0u;
    const bool unsafePaste = ! runtime.pasteIsSafe(pasteSource.data(), 3u);
    const GhosttyResult pasteResult = runtime.pasteEncode(
        pasteSource.data(), 3u, true, pasteOutput.data(), pasteOutput.size(), &pasteLength);
    check(unsafePaste);
    check(pasteResult == GHOSTTY_SUCCESS && pasteLength == 15u &&
          std::string_view(pasteOutput.data(), pasteLength) == "\x1b[200~a\nb\x1b[201~");

    GhosttySelection selection{};
    selection.size = sizeof(selection);
    const bool selectedAll = runtime.terminalSelectAll(terminal, &selection) == GHOSTTY_SUCCESS &&
        runtime.terminalSet(terminal, GHOSTTY_TERMINAL_OPT_SELECTION, &selection) == GHOSTTY_SUCCESS;
    check(selectedAll);
    GhosttyTerminalSelectionFormatOptions selectionOptions{};
    selectionOptions.size = sizeof(selectionOptions);
    selectionOptions.emit = GHOSTTY_FORMATTER_FORMAT_PLAIN;
    selectionOptions.unwrap = true;
    selectionOptions.trim = true;
    std::array<uint8_t, 64u> selectedText{};
    size_t selectedLength = 0u;
    check(selectedAll && runtime.terminalSelectionFormatBuffer(
                             terminal,
                             selectionOptions,
                             selectedText.data(),
                             selectedText.size(),
                             &selectedLength) == GHOSTTY_SUCCESS &&
          selectedLength != 0u && std::string_view(reinterpret_cast<const char*>(selectedText.data()), selectedLength).contains('A'));

    GhosttyPoint trackedPoint{};
    trackedPoint.tag = GHOSTTY_POINT_TAG_VIEWPORT;
    trackedPoint.value.coordinate.x = 0u;
    trackedPoint.value.coordinate.y = 0u;
    GhosttyGridRef trackedSnapshot{};
    trackedSnapshot.size = sizeof(trackedSnapshot);
    const bool trackedSelectionAnchor =
        runtime.terminalGridRefTrack(terminal, trackedPoint, &trackedCell) == GHOSTTY_SUCCESS && trackedCell != nullptr &&
        runtime.trackedGridRefSnapshot(trackedCell, &trackedSnapshot) == GHOSTTY_SUCCESS && trackedSnapshot.node != nullptr;
    check(trackedSelectionAnchor);

    constexpr std::string_view mouseMode = "\x1b[?1000h\x1b[?1006h";
    runtime.terminalVtWrite(terminal, reinterpret_cast<const uint8_t*>(mouseMode.data()), mouseMode.size());
    runtime.mouseEncoderSetFromTerminal(mouseEncoder, terminal);
    GhosttyMouseEncoderSize mouseSize{};
    mouseSize.size = sizeof(mouseSize);
    mouseSize.screen_width = 80u;
    mouseSize.screen_height = 40u;
    mouseSize.cell_width = 10u;
    mouseSize.cell_height = 20u;
    runtime.mouseEncoderSetOption(mouseEncoder, GHOSTTY_MOUSE_ENCODER_OPT_SIZE, &mouseSize);
    const auto appendMouseBytes = [&](GhosttyMouseAction action,
                                      GhosttyMouseButton button,
                                      bool anyButtonPressed,
                                      std::string& output) noexcept
    {
        runtime.mouseEncoderSetOption(
            mouseEncoder, GHOSTTY_MOUSE_ENCODER_OPT_ANY_BUTTON_PRESSED, &anyButtonPressed);
        runtime.mouseEventSetAction(mouseEvent, action);
        runtime.mouseEventSetButton(mouseEvent, button);
        runtime.mouseEventSetMods(mouseEvent, 0u);
        runtime.mouseEventSetPosition(mouseEvent, GhosttyMousePosition{5.0f, 5.0f});
        std::array<char, 64u> encoded{};
        size_t encodedLength = 0u;
        const GhosttyResult result = runtime.mouseEncoderEncode(
            mouseEncoder, mouseEvent, encoded.data(), encoded.size(), &encodedLength);
        if (result != GHOSTTY_SUCCESS || encodedLength == 0u)
        {
            return false;
        }
        output.append(encoded.data(), encodedLength);
        return true;
    };
    const auto pressMouseButton = [&](uint8_t& buttons, GhosttyMouseButton button, std::string& output) noexcept
    {
        return AddReportedMouseButton(buttons, button) &&
            appendMouseBytes(GHOSTTY_MOUSE_ACTION_PRESS, button, buttons != 0u, output);
    };
    const auto releaseMouseButton = [&](uint8_t& buttons, GhosttyMouseButton button, std::string& output) noexcept
    {
        return RemoveReportedMouseButton(buttons, button) &&
            appendMouseBytes(GHOSTTY_MOUSE_ACTION_RELEASE, button, buttons != 0u, output);
    };

    uint8_t reportedButtons = 0u;
    std::string singleMouseBytes;
    const bool singleMouse = pressMouseButton(reportedButtons, GHOSTTY_MOUSE_BUTTON_LEFT, singleMouseBytes) &&
        releaseMouseButton(reportedButtons, GHOSTTY_MOUSE_BUTTON_LEFT, singleMouseBytes);
    check(singleMouse && reportedButtons == 0u && singleMouseBytes == "\x1b[<0;1;1M\x1b[<0;1;1m");

    std::string chordMouseBytes;
    const bool chordMouse = pressMouseButton(reportedButtons, GHOSTTY_MOUSE_BUTTON_LEFT, chordMouseBytes) &&
        pressMouseButton(reportedButtons, GHOSTTY_MOUSE_BUTTON_RIGHT, chordMouseBytes) &&
        releaseMouseButton(reportedButtons, GHOSTTY_MOUSE_BUTTON_RIGHT, chordMouseBytes) &&
        releaseMouseButton(reportedButtons, GHOSTTY_MOUSE_BUTTON_LEFT, chordMouseBytes);
    check(chordMouse && reportedButtons == 0u &&
          chordMouseBytes == "\x1b[<0;1;1M\x1b[<2;1;1M\x1b[<2;1;1m\x1b[<0;1;1m");

    std::string mismatchedMouseBytes;
    const bool mismatchedMousePress = pressMouseButton(reportedButtons, GHOSTTY_MOUSE_BUTTON_LEFT, mismatchedMouseBytes);
    const size_t mismatchedBytesBefore = mismatchedMouseBytes.size();
    const bool mismatchedReleaseIgnored = ! RemoveReportedMouseButton(reportedButtons, GHOSTTY_MOUSE_BUTTON_RIGHT) &&
        mismatchedMouseBytes.size() == mismatchedBytesBefore &&
        HasReportedMouseButton(reportedButtons, GHOSTTY_MOUSE_BUTTON_LEFT);
    const bool mismatchedMouseCleanup = releaseMouseButton(
        reportedButtons, GHOSTTY_MOUSE_BUTTON_LEFT, mismatchedMouseBytes);
    check(mismatchedMousePress && mismatchedReleaseIgnored && mismatchedMouseCleanup && reportedButtons == 0u &&
          mismatchedMouseBytes == "\x1b[<0;1;1M\x1b[<0;1;1m");

    std::string captureLossMouseBytes;
    bool captureLossMouse = pressMouseButton(reportedButtons, GHOSTTY_MOUSE_BUTTON_LEFT, captureLossMouseBytes) &&
        pressMouseButton(reportedButtons, GHOSTTY_MOUSE_BUTTON_MIDDLE, captureLossMouseBytes);
    for (const GhosttyMouseButton button : kTrackedMouseButtons)
    {
        if (HasReportedMouseButton(reportedButtons, button))
        {
            captureLossMouse = releaseMouseButton(reportedButtons, button, captureLossMouseBytes) && captureLossMouse;
        }
    }
    check(captureLossMouse && reportedButtons == 0u &&
          captureLossMouseBytes == "\x1b[<0;1;1M\x1b[<1;1;1M\x1b[<0;1;1m\x1b[<1;1;1m");

    const uint64_t kittyStorageLimit = kKittyImageStorageBytes;
    const bool kittyExternalMedium = false;
    const size_t kittyApcLimit = kMaximumInputBytes;
    const GhosttyKittyPlacementLimits kittyPlacementLimits{
        sizeof(GhosttyKittyPlacementLimits), kMaximumKittyPlacements, kMaximumKittyPlacementBytes};
    const bool kittyConfigured =
        runtime.terminalSet(terminal, GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_STORAGE_LIMIT, &kittyStorageLimit) == GHOSTTY_SUCCESS &&
        runtime.terminalSet(terminal, GHOSTTY_TERMINAL_OPT_KITTY_PLACEMENT_LIMITS, &kittyPlacementLimits) == GHOSTTY_SUCCESS &&
        runtime.terminalSet(terminal, GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_FILE, &kittyExternalMedium) == GHOSTTY_SUCCESS &&
        runtime.terminalSet(terminal, GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_TEMP_FILE, nullptr) == GHOSTTY_SUCCESS &&
        runtime.terminalSet(terminal, GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_SHARED_MEM, &kittyExternalMedium) == GHOSTTY_SUCCESS &&
        runtime.terminalSet(terminal, GHOSTTY_TERMINAL_OPT_APC_MAX_BYTES_KITTY, &kittyApcLimit) == GHOSTTY_SUCCESS;
    check(kittyConfigured);
    constexpr std::string_view kittyPixel = "\x1b_Ga=T,f=24,s=1,v=1,q=2;/wAA\x1b\\";
    runtime.terminalVtWrite(terminal, reinterpret_cast<const uint8_t*>(kittyPixel.data()), kittyPixel.size());
    GhosttyKittyGraphics graphics = nullptr;
    bool kittyPixelReady = kittyConfigured &&
        runtime.terminalGet(terminal, GHOSTTY_TERMINAL_DATA_KITTY_GRAPHICS, &graphics) == GHOSTTY_SUCCESS && graphics != nullptr &&
        runtime.kittyGraphicsGet(graphics, GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_ITERATOR, &kittyPlacements) == GHOSTTY_SUCCESS &&
        runtime.kittyPlacementNext(kittyPlacements);
    if (kittyPixelReady)
    {
        uint32_t imageId = 0u;
        kittyPixelReady = runtime.kittyPlacementGet(
                              kittyPlacements, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_IMAGE_ID, &imageId) == GHOSTTY_SUCCESS;
        const GhosttyKittyGraphicsImage image = kittyPixelReady ? runtime.kittyGraphicsImage(graphics, imageId) : nullptr;
        const uint8_t* data = nullptr;
        size_t dataLength = 0u;
        GhosttyKittyImageFormat format = GHOSTTY_KITTY_IMAGE_FORMAT_MAX_VALUE;
        const std::array<GhosttyKittyGraphicsImageData, 3u> keys{
            GHOSTTY_KITTY_IMAGE_DATA_FORMAT, GHOSTTY_KITTY_IMAGE_DATA_DATA_PTR, GHOSTTY_KITTY_IMAGE_DATA_DATA_LEN};
        std::array<void*, 3u> values{&format, &data, &dataLength};
        kittyPixelReady = image != nullptr &&
            runtime.kittyGraphicsImageGetMulti(image, keys.size(), keys.data(), values.data(), nullptr) == GHOSTTY_SUCCESS &&
            format == GHOSTTY_KITTY_IMAGE_FORMAT_RGB && data != nullptr && dataLength == 3u && data[0] == 0xFFu &&
            data[1] == 0u && data[2] == 0u;
    }
    check(kittyPixelReady);

    const bool placementBoundCreated =
        runtime.terminalNew(nullptr, &placementBoundTerminal, 8u, 2u) == GHOSTTY_SUCCESS &&
        placementBoundTerminal != nullptr &&
        runtime.terminalSet(
            placementBoundTerminal, GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_STORAGE_LIMIT, &kittyStorageLimit) == GHOSTTY_SUCCESS &&
        runtime.terminalSet(
            placementBoundTerminal, GHOSTTY_TERMINAL_OPT_KITTY_PLACEMENT_LIMITS, &kittyPlacementLimits) == GHOSTTY_SUCCESS &&
        runtime.terminalSet(
            placementBoundTerminal, GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_FILE, &kittyExternalMedium) == GHOSTTY_SUCCESS &&
        runtime.terminalSet(
            placementBoundTerminal, GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_TEMP_FILE, nullptr) == GHOSTTY_SUCCESS &&
        runtime.terminalSet(
            placementBoundTerminal, GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_SHARED_MEM, &kittyExternalMedium) == GHOSTTY_SUCCESS &&
        runtime.terminalSet(
            placementBoundTerminal, GHOSTTY_TERMINAL_OPT_APC_MAX_BYTES_KITTY, &kittyApcLimit) == GHOSTTY_SUCCESS;
    constexpr std::string_view boundedImage = "\x1b_Ga=t,t=d,f=24,i=901,s=1,v=1,q=2;/wAA\x1b\\";
    if (placementBoundCreated)
    {
        runtime.terminalVtWrite(
            placementBoundTerminal, reinterpret_cast<const uint8_t*>(boundedImage.data()), boundedImage.size());
        for (uint32_t placementId = 1u; placementId <= kMaximumKittyPlacements; ++placementId)
        {
            const std::string display = std::format("\x1b_Ga=p,i=901,p={},c=1,r=1,q=2;\x1b\\", placementId);
            runtime.terminalVtWrite(
                placementBoundTerminal, reinterpret_cast<const uint8_t*>(display.data()), display.size());
        }
    }
    GhosttyKittyGraphics boundedGraphics = nullptr;
    size_t retainedPlacementCountBeforeRejection = 0u;
    size_t retainedPlacementBytesBeforeRejection = 0u;
    size_t retainedPlacementCountLimit = 0u;
    size_t retainedPlacementBytesLimit = 0u;
    const bool placementAccountingReady = placementBoundCreated &&
        runtime.terminalGet(
            placementBoundTerminal, GHOSTTY_TERMINAL_DATA_KITTY_GRAPHICS, &boundedGraphics) == GHOSTTY_SUCCESS &&
        boundedGraphics != nullptr &&
        runtime.kittyGraphicsGet(
            boundedGraphics,
            GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_COUNT,
            &retainedPlacementCountBeforeRejection) == GHOSTTY_SUCCESS &&
        runtime.kittyGraphicsGet(
            boundedGraphics,
            GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_ALLOCATED_BYTES,
            &retainedPlacementBytesBeforeRejection) == GHOSTTY_SUCCESS &&
        runtime.kittyGraphicsGet(
            boundedGraphics,
            GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_COUNT_LIMIT,
            &retainedPlacementCountLimit) == GHOSTTY_SUCCESS &&
        runtime.kittyGraphicsGet(
            boundedGraphics,
            GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_BYTES_LIMIT,
            &retainedPlacementBytesLimit) == GHOSTTY_SUCCESS;
    if (placementAccountingReady)
    {
        constexpr std::string_view replacement = "\x1b_Ga=p,i=901,p=1,c=1,r=1,q=2;\x1b\\";
        constexpr std::string_view overLimit = "\x1b_Ga=p,i=901,p=4097,c=1,r=1,q=2;\x1b\\";
        runtime.terminalVtWrite(
            placementBoundTerminal, reinterpret_cast<const uint8_t*>(replacement.data()), replacement.size());
        runtime.terminalVtWrite(
            placementBoundTerminal, reinterpret_cast<const uint8_t*>(overLimit.data()), overLimit.size());
    }
    boundedGraphics = nullptr;
    size_t retainedPlacementCountAfterRejection = 0u;
    size_t retainedPlacementBytesAfterRejection = 0u;
    const bool placementRejectionReady = placementAccountingReady &&
        runtime.terminalGet(
            placementBoundTerminal, GHOSTTY_TERMINAL_DATA_KITTY_GRAPHICS, &boundedGraphics) == GHOSTTY_SUCCESS &&
        boundedGraphics != nullptr &&
        runtime.kittyGraphicsGet(
            boundedGraphics,
            GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_COUNT,
            &retainedPlacementCountAfterRejection) == GHOSTTY_SUCCESS &&
        runtime.kittyGraphicsGet(
            boundedGraphics,
            GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_ALLOCATED_BYTES,
            &retainedPlacementBytesAfterRejection) == GHOSTTY_SUCCESS;
    size_t retainedPlacementCount = 0u;
    const bool placementIteratorReady = placementRejectionReady &&
        runtime.kittyGraphicsGet(
            boundedGraphics,
            GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_ITERATOR,
            &kittyPlacements) == GHOSTTY_SUCCESS;
    while (placementIteratorReady && runtime.kittyPlacementNext(kittyPlacements))
    {
        ++retainedPlacementCount;
    }
    check(placementIteratorReady && retainedPlacementCount == kMaximumKittyPlacements &&
          retainedPlacementCountBeforeRejection == kMaximumKittyPlacements &&
          retainedPlacementCountAfterRejection == retainedPlacementCountBeforeRejection &&
          retainedPlacementBytesBeforeRejection > 0u &&
          retainedPlacementBytesBeforeRejection <= kMaximumKittyPlacementBytes &&
          retainedPlacementBytesAfterRejection == retainedPlacementBytesBeforeRejection &&
          retainedPlacementCountLimit == kMaximumKittyPlacements &&
          retainedPlacementBytesLimit == kMaximumKittyPlacementBytes);

    constexpr std::string_view kittyPng =
        "\x1b_Ga=T,f=100,q=2;"
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAA"
        "DUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg=="
        "\x1b\\";
    runtime.terminalVtWrite(terminal, reinterpret_cast<const uint8_t*>(kittyPng.data()), kittyPng.size());
    bool kittyPngReady = runtime.terminalGet(
                             terminal, GHOSTTY_TERMINAL_DATA_KITTY_GRAPHICS, &graphics) == GHOSTTY_SUCCESS &&
        graphics != nullptr &&
        runtime.kittyGraphicsGet(
            graphics, GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_ITERATOR, &kittyPlacements) == GHOSTTY_SUCCESS;
    bool foundDecodedPng = false;
    while (kittyPngReady && runtime.kittyPlacementNext(kittyPlacements))
    {
        uint32_t imageId = 0u;
        if (runtime.kittyPlacementGet(
                kittyPlacements, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_IMAGE_ID, &imageId) != GHOSTTY_SUCCESS)
        {
            continue;
        }
        const GhosttyKittyGraphicsImage image = runtime.kittyGraphicsImage(graphics, imageId);
        uint32_t width = 0u;
        uint32_t height = 0u;
        GhosttyKittyImageFormat format = GHOSTTY_KITTY_IMAGE_FORMAT_MAX_VALUE;
        size_t dataLength = 0u;
        const std::array<GhosttyKittyGraphicsImageData, 4u> keys{GHOSTTY_KITTY_IMAGE_DATA_WIDTH,
                                                                 GHOSTTY_KITTY_IMAGE_DATA_HEIGHT,
                                                                 GHOSTTY_KITTY_IMAGE_DATA_FORMAT,
                                                                 GHOSTTY_KITTY_IMAGE_DATA_DATA_LEN};
        std::array<void*, 4u> values{&width, &height, &format, &dataLength};
        if (image != nullptr &&
            runtime.kittyGraphicsImageGetMulti(
                image, keys.size(), keys.data(), values.data(), nullptr) == GHOSTTY_SUCCESS &&
            width == 1u && height == 1u && format == GHOSTTY_KITTY_IMAGE_FORMAT_RGBA && dataLength == 4u)
        {
            foundDecodedPng = true;
            break;
        }
    }
    check(kittyPngReady && foundDecodedPng);

    constexpr std::string_view scrollbackText =
        "\r\n0\r\n1\r\n2\r\n3\r\n4\r\n5\r\n6\r\n7\r\n8\r\n9\r\n10\r\n11\r\n12\r\n13\r\n14\r\n15";
    runtime.terminalVtWrite(
        terminal, reinterpret_cast<const uint8_t*>(scrollbackText.data()), scrollbackText.size());
    GhosttyTerminalScrollbar scrollbar{};
    const bool scrollbackReady =
        runtime.terminalGet(terminal, GHOSTTY_TERMINAL_DATA_SCROLLBAR, &scrollbar) == GHOSTTY_SUCCESS &&
        scrollbar.total > scrollbar.len && scrollbar.offset == scrollbar.total - scrollbar.len;
    check(scrollbackReady);

    GhosttyTerminalScrollViewport scroll{};
    scroll.tag = GHOSTTY_SCROLL_VIEWPORT_ROW;
    scroll.value.row = 0u;
    runtime.terminalScrollViewport(terminal, scroll);
    check(scrollbackReady &&
          runtime.terminalGet(terminal, GHOSTTY_TERMINAL_DATA_SCROLLBAR, &scrollbar) == GHOSTTY_SUCCESS &&
          scrollbar.offset == 0u);

    const uint64_t maximumOffset = scrollbar.total - scrollbar.len;
    scroll.value.row = static_cast<size_t>(maximumOffset);
    runtime.terminalScrollViewport(terminal, scroll);
    check(runtime.terminalGet(terminal, GHOSTTY_TERMINAL_DATA_SCROLLBAR, &scrollbar) == GHOSTTY_SUCCESS &&
          scrollbar.offset == maximumOffset);

    const auto makeMaximumKittyWork = [](uint64_t requestId) -> TerminalKittyGenerationWork
    {
        TerminalKittyGenerationWork maximumWork{};
        maximumWork.requestId = requestId;
        maximumWork.storageGeneration = requestId;
        TerminalKittyImageWork maximumImage{};
        maximumImage.imageId = 1u;
        maximumImage.width = 4096u;
        maximumImage.height = 4096u;
        maximumImage.imageGeneration = requestId;
        maximumImage.format = TerminalKittyPixelFormat::Rgba;
        maximumImage.source.resize(TerminalKittyImagePipeline::MaximumSourceBytes);
        maximumImage.source[0] = 0xFFu;
        maximumImage.source[3] = 0xFFu;
        maximumImage.source[maximumImage.source.size() - 1u] = 0xFFu;
        maximumWork.images.push_back(std::move(maximumImage));
        return maximumWork;
    };

    {
        TerminalKittyGenerationWork maximumWork = makeMaximumKittyWork(1u);
        std::atomic_uint32_t maximumNotifications{0u};
        TerminalKittyImagePipeline maximumPipeline;
        const bool maximumSubmitted =
            maximumPipeline.Start(&CountKittyPipelineNotification, &maximumNotifications) == S_OK &&
            maximumPipeline.Submit(std::move(maximumWork));
        std::optional<TerminalKittyGenerationResult> maximumResult;
        const ULONGLONG maximumDeadline = GetTickCount64() + 10000u;
        while (maximumSubmitted && GetTickCount64() < maximumDeadline && ! maximumResult.has_value())
        {
            maximumResult = maximumPipeline.TakeReady(1u);
            if (! maximumResult.has_value())
            {
                Sleep(1u);
            }
        }
        const TerminalKittyPipelineStats maximumStats = maximumPipeline.GetStats();
        maximumPipeline.Stop();
        check(maximumResult.has_value() &&
              maximumResult->convertedBytes == TerminalKittyImagePipeline::MaximumConvertedBytes &&
              maximumResult->images.size() == 1u &&
              maximumResult->images[0].bgra.size() == TerminalKittyImagePipeline::MaximumConvertedBytes &&
              maximumResult->images[0].bgra[2] == 0xFFu && maximumResult->images[0].bgra[3] == 0xFFu &&
              maximumStats.memoryHighWaterBytes == TerminalKittyImagePipeline::MaximumSourceBytes +
                      TerminalKittyImagePipeline::MaximumConvertedBytes &&
               maximumStats.maximumActiveWorkers == 1u && maximumNotifications.load(std::memory_order_acquire) == 1u);
    }

    {
        std::atomic_uint32_t activeCloseNotifications{0u};
        TerminalKittyImagePipeline activeClosePipeline;
        const bool activeCloseStarted =
            activeClosePipeline.Start(&CountKittyPipelineNotification, &activeCloseNotifications) == S_OK;
        activeClosePipeline.SetPauseAfterConversionStartForTests(true);
        const bool activeCloseSubmitted =
            activeCloseStarted && activeClosePipeline.Submit(makeMaximumKittyWork(1u));
        const bool conversionStarted =
            activeCloseSubmitted && activeClosePipeline.WaitForConversionStartForTests(5000u);
        const auto activeCloseStartedAt = std::chrono::steady_clock::now();
        activeClosePipeline.Stop();
        const uint64_t activeCloseUs = Debug::Perf::ElapsedUs(activeCloseStartedAt);
        const TerminalKittyPipelineStats activeCloseStats = activeClosePipeline.GetStats();
        Debug::Perf::EmitDurationUs(L"terminal.kitty.active_close_us",
                                    std::max<uint64_t>(1u, activeCloseUs),
                                    TerminalKittyImagePipeline::MaximumSourceBytes,
                                    TerminalKittyImagePipeline::MaximumConvertedBytes);
        check(conversionStarted && activeCloseUs < 5'000'000u && activeCloseStats.activeWorkers == 0u &&
              activeCloseStats.queuedBytes == 0u && activeCloseStats.activeSourceBytes == 0u &&
              activeCloseStats.activeConvertedBytes == 0u && activeCloseStats.readyBytes == 0u &&
              activeCloseNotifications.load(std::memory_order_acquire) == 0u);
    }

    const auto makeKittyWork = [](uint64_t requestId, uint8_t red) -> TerminalKittyGenerationWork
    {
        TerminalKittyGenerationWork work{};
        work.requestId = requestId;
        work.storageGeneration = requestId;
        TerminalKittyImageWork image{};
        image.imageId = 1u;
        image.width = 1u;
        image.height = 1u;
        image.imageGeneration = requestId;
        image.format = TerminalKittyPixelFormat::Rgba;
        image.source = {red, 0u, 0u, 0xFFu};
        work.images.push_back(std::move(image));
        return work;
    };
    std::atomic_uint32_t kittyNotifications{0u};
    TerminalKittyImagePipeline kittyPipeline;
    check(kittyPipeline.Start(&CountKittyPipelineNotification, &kittyNotifications) == S_OK);
    kittyPipeline.SetPauseBeforeTakeForTests(true);
    check(kittyPipeline.Submit(makeKittyWork(1u, 0x11u)));
    check(kittyPipeline.Submit(makeKittyWork(2u, 0x22u)));
    check(kittyPipeline.Submit(makeKittyWork(3u, 0x33u)));
    TerminalKittyPipelineStats kittyStats = kittyPipeline.GetStats();
    check(kittyStats.queuedBytes == 4u && kittyStats.latestRequestId == 3u &&
          kittyStats.droppedPendingGenerations == 2u && kittyStats.maximumActiveWorkers <= 1u);
    kittyPipeline.SetPauseBeforeTakeForTests(false);

    const auto waitForReady = [&kittyPipeline](uint64_t requestId) -> std::optional<TerminalKittyGenerationResult>
    {
        const ULONGLONG deadline = GetTickCount64() + 5000u;
        while (GetTickCount64() < deadline)
        {
            std::optional<TerminalKittyGenerationResult> result = kittyPipeline.TakeReady(requestId);
            if (result.has_value())
            {
                return result;
            }
            Sleep(1u);
        }
        return std::nullopt;
    };
    std::optional<TerminalKittyGenerationResult> replacement = waitForReady(3u);
    check(replacement.has_value() && replacement->images.size() == 1u && replacement->images[0].bgra[2] == 0x33u);

    kittyPipeline.SetPauseBeforePublishForTests(true);
    check(kittyPipeline.Submit(makeKittyWork(4u, 0x44u)));
    const ULONGLONG activeDeadline = GetTickCount64() + 5000u;
    while (GetTickCount64() < activeDeadline && kittyPipeline.GetStats().activeRequestId != 4u)
    {
        Sleep(1u);
    }
    check(kittyPipeline.GetStats().activeRequestId == 4u);
    check(kittyPipeline.Submit(makeKittyWork(5u, 0x55u)));
    kittyPipeline.SetPauseBeforePublishForTests(false);
    std::optional<TerminalKittyGenerationResult> newest = waitForReady(5u);
    kittyStats = kittyPipeline.GetStats();
    check(newest.has_value() && newest->images.size() == 1u && newest->images[0].bgra[2] == 0x55u &&
          kittyStats.rejectedStaleGenerations >= 1u && kittyStats.maximumActiveWorkers == 1u &&
          kittyStats.memoryHighWaterBytes <= TerminalKittyImagePipeline::MaximumSourceBytes +
                  TerminalKittyImagePipeline::MaximumConvertedBytes * 2u);

    TerminalKittyGenerationWork oversizedWork{};
    oversizedWork.requestId = 6u;
    oversizedWork.storageGeneration = 6u;
    TerminalKittyImageWork oversizedImage{};
    oversizedImage.imageId = 1u;
    oversizedImage.width = 4097u;
    oversizedImage.height = 4096u;
    oversizedImage.imageGeneration = 6u;
    oversizedImage.format = TerminalKittyPixelFormat::Rgba;
    oversizedWork.images.push_back(std::move(oversizedImage));
    check(! kittyPipeline.Submit(std::move(oversizedWork)));

    kittyPipeline.SetPauseBeforeTakeForTests(true);
    check(kittyPipeline.Submit(makeKittyWork(6u, 0x66u)));
    const ULONGLONG stopStartedAt = GetTickCount64();
    kittyPipeline.Stop();
    kittyStats = kittyPipeline.GetStats();
    check(GetTickCount64() - stopStartedAt < 2000u && kittyStats.activeWorkers == 0u &&
          kittyStats.queuedBytes == 0u && kittyNotifications.load(std::memory_order_acquire) >= 2u);

    {
        Terminal hostKittyTerminal;
        static_cast<void>(hostKittyTerminal.DebugRunKittyCacheSelfTests(passedTests, failedTests));
    }

    return *failedTests == 0u ? S_OK : E_FAIL;
}
#endif

Terminal::Terminal()
{
    g_terminalObjectCount.fetch_add(1u, std::memory_order_acq_rel);
    _metaId = L"builtin/terminal";
    _metaShortId = L"terminal";
    _metaName = LoadStringResource(g_hInstance, IDS_TERMINAL_NAME);
    _metaDescription = LoadStringResource(g_hInstance, IDS_TERMINAL_DESCRIPTION);
    _unsafePasteTitle = LoadStringResource(g_hInstance, IDS_TERMINAL_UNSAFE_PASTE_TITLE);
    _unsafePasteMessage = LoadStringResource(g_hInstance, IDS_TERMINAL_UNSAFE_PASTE_MESSAGE);
    _unsafePasteAccept = LoadStringResource(g_hInstance, IDS_TERMINAL_UNSAFE_PASTE_ACCEPT);
    _unsafePasteCancel = LoadStringResource(g_hInstance, IDS_TERMINAL_UNSAFE_PASTE_CANCEL);
    _hyperlinkTitle = LoadStringResource(g_hInstance, IDS_TERMINAL_HYPERLINK_TITLE);
    _hyperlinkMessage = LoadStringResource(g_hInstance, IDS_TERMINAL_HYPERLINK_MESSAGE);
    _hyperlinkAccept = LoadStringResource(g_hInstance, IDS_TERMINAL_HYPERLINK_ACCEPT);
    _hyperlinkCancel = LoadStringResource(g_hInstance, IDS_TERMINAL_HYPERLINK_CANCEL);
    _osc52Title = LoadStringResource(g_hInstance, IDS_TERMINAL_OSC52_TITLE);
    _osc52Message = LoadStringResource(g_hInstance, IDS_TERMINAL_OSC52_MESSAGE);
    _osc52Accept = LoadStringResource(g_hInstance, IDS_TERMINAL_OSC52_ACCEPT);
    _osc52Cancel = LoadStringResource(g_hInstance, IDS_TERMINAL_OSC52_CANCEL);
    _metaData.id = _metaId.c_str();
    _metaData.shortId = _metaShortId.c_str();
    _metaData.name = _metaName.c_str();
    _metaData.description = _metaDescription.c_str();
    _metaData.author = L"RedSalamander";
    _metaData.version = VERSINFO_PLUGIN_VERSION;
    static_cast<void>(SetConfiguration(nullptr));
    // Close must stay nonblocking even when Open never ran (ABI contract tests
    // and last-Release retirement). Create the pinned close work here so the
    // UI-facing Close path never falls through to finishClose.
    static_cast<void>(ensureCloseWork());
}

Terminal::~Terminal()
{
    static_cast<void>(SetCallback(nullptr, nullptr));
    if (! _closing.exchange(true, std::memory_order_acq_rel))
    {
        // Close work is created in the constructor. Reaching the destructor
        // without Close means last-Release could not submit the callback (work
        // creation failed) or a stack-allocated debug fixture is tearing down.
        // There is no live session to join on the UI thread in that case.
        prepareClose();
        finishClose();
    }
    g_terminalObjectCount.fetch_sub(1u, std::memory_order_acq_rel);
}

HRESULT STDMETHODCALLTYPE Terminal::QueryInterface(REFIID riid, void** object) noexcept
{
    if (object == nullptr)
    {
        return E_POINTER;
    }
    *object = nullptr;
    if (riid == __uuidof(IUnknown) || riid == __uuidof(ITerminal))
    {
        *object = static_cast<ITerminal*>(this);
    }
    else if (riid == __uuidof(ITerminalActions))
    {
        *object = static_cast<ITerminalActions*>(this);
    }
    else if (riid == __uuidof(IInformations))
    {
        *object = static_cast<IInformations*>(this);
    }
    else
    {
        return E_NOINTERFACE;
    }
    AddRef();
    return S_OK;
}

ULONG STDMETHODCALLTYPE Terminal::AddRef() noexcept
{
    return _refCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
}

ULONG STDMETHODCALLTYPE Terminal::Release() noexcept
{
    const ULONG remaining = _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
    if (remaining == 0u)
    {
        if (! _closing.load(std::memory_order_acquire) && _closeWork)
        {
            // A host is required to call Close before its final Release, but a
            // defensive host-retirement path must not move ConPTY teardown back
            // onto the releasing (normally UI) thread. Revive one private
            // retirement reference, let Close add the callback reference, and
            // drop the private reference without recursively entering Release.
            _refCount.store(1u, std::memory_order_release);
            static_cast<void>(Close());
            const ULONG callbackReferences = _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
            if (callbackReferences != 0u)
            {
                return 0u;
            }
        }
        delete this;
    }
    return remaining;
}

HRESULT STDMETHODCALLTYPE Terminal::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (metaData == nullptr)
    {
        return E_POINTER;
    }
    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE Terminal::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }
    *schemaJsonUtf8 = GetTerminalStaticConfigurationSchema();
    return *schemaJsonUtf8 == nullptr ? E_OUTOFMEMORY : S_OK;
}

HRESULT STDMETHODCALLTYPE Terminal::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    Config next;
    if (configurationJsonUtf8 != nullptr && configurationJsonUtf8[0] != '\0')
    {
        const std::string_view json(configurationJsonUtf8);
        yyjson_doc* document = yyjson_read(json.data(), json.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
        if (document == nullptr)
        {
            return E_INVALIDARG;
        }
        const auto freeDocument = wil::scope_exit([&] noexcept { yyjson_doc_free(document); });
        yyjson_val* root = yyjson_doc_get_root(document);
        if (root == nullptr || ! yyjson_is_obj(root))
        {
            return E_INVALIDARG;
        }

        if (yyjson_val* shell = yyjson_obj_get(root, "defaultShell"); shell != nullptr)
        {
            if (! yyjson_is_str(shell))
            {
                return E_INVALIDARG;
            }
            const std::wstring value = Common::Strings::Utf16FromUtf8StrictOrEmpty(yyjson_get_str(shell));
            if (value != L"auto" && value != L"powershell" && value != L"cmd")
            {
                return E_INVALIDARG;
            }
            next.defaultShell = value;
        }
        if (yyjson_val* family = yyjson_obj_get(root, "fontFamily"); family != nullptr)
        {
            if (! yyjson_is_str(family))
            {
                return E_INVALIDARG;
            }
            next.fontFamily = Common::Strings::Utf16FromUtf8StrictOrEmpty(yyjson_get_str(family));
            if (next.fontFamily.empty() || next.fontFamily.size() > 128u)
            {
                return E_INVALIDARG;
            }
        }
        if (yyjson_val* size = yyjson_obj_get(root, "fontSizeDip"); size != nullptr)
        {
            if (! yyjson_is_num(size))
            {
                return E_INVALIDARG;
            }
            const double value = yyjson_get_real(size);
            if (! std::isfinite(value) || value < 8.0 || value > 32.0)
            {
                return E_INVALIDARG;
            }
            next.fontSizeDip = static_cast<float>(value);
        }
        if (yyjson_val* maxMiB = yyjson_obj_get(root, "maxFormattedMiB"); maxMiB != nullptr)
        {
            if (! yyjson_is_int(maxMiB))
            {
                return E_INVALIDARG;
            }
            const int64_t value = yyjson_get_int(maxMiB);
            if (value < 1 || value > 32)
            {
                return E_INVALIDARG;
            }
            next.maxFormattedBytes = static_cast<uint32_t>(value) * 1024u * 1024u;
        }
        if (yyjson_val* pasteMaxBytes = yyjson_obj_get(root, "pasteMaxBytes"); pasteMaxBytes != nullptr)
        {
            if (! yyjson_is_int(pasteMaxBytes))
            {
                return E_INVALIDARG;
            }
            const int64_t value = yyjson_get_int(pasteMaxBytes);
            if (value < 4096 || value > static_cast<int64_t>(kMaximumInputBytes))
            {
                return E_INVALIDARG;
            }
            next.pasteMaxBytes = static_cast<uint32_t>(value);
        }
        if (! ReadBoolean(yyjson_obj_get(root, "warnOnUnsafePaste"), next.warnOnUnsafePaste))
        {
            return E_INVALIDARG;
        }
        if (! ReadBoolean(yyjson_obj_get(root, "followPathWhenIdle"), next.followPathWhenIdle))
        {
            return E_INVALIDARG;
        }
        if (yyjson_val* hyperlinkPolicy = yyjson_obj_get(root, "hyperlinkPolicy"); hyperlinkPolicy != nullptr)
        {
            if (! yyjson_is_str(hyperlinkPolicy))
            {
                return E_INVALIDARG;
            }
            const std::wstring value = Common::Strings::Utf16FromUtf8StrictOrEmpty(yyjson_get_str(hyperlinkPolicy));
            if (value != L"disabled" && value != L"ask" && value != L"open")
            {
                return E_INVALIDARG;
            }
            next.hyperlinkPolicy = value;
        }
        if (yyjson_val* osc52Policy = yyjson_obj_get(root, "osc52Policy"); osc52Policy != nullptr)
        {
            if (! yyjson_is_str(osc52Policy))
            {
                return E_INVALIDARG;
            }
            const std::wstring value = Common::Strings::Utf16FromUtf8StrictOrEmpty(yyjson_get_str(osc52Policy));
            if (value != L"deny" && value != L"ask" && value != L"allow")
            {
                return E_INVALIDARG;
            }
            next.osc52Policy = value;
        }
        if (yyjson_val* osc52MaxBytes = yyjson_obj_get(root, "osc52MaxBytes"); osc52MaxBytes != nullptr)
        {
            if (! yyjson_is_int(osc52MaxBytes))
            {
                return E_INVALIDARG;
            }
            const int64_t value = yyjson_get_int(osc52MaxBytes);
            if (value < kMinimumOsc52Bytes || value > kMaximumOsc52Bytes)
            {
                return E_INVALIDARG;
            }
            next.osc52MaxBytes = static_cast<uint32_t>(value);
        }
    }

    const std::string fontUtf8 = Common::Strings::Utf8FromUtf16StrictOrEmpty(next.fontFamily);
    yyjson_mut_doc* outputDocument = yyjson_mut_doc_new(nullptr);
    if (outputDocument == nullptr)
    {
        return E_OUTOFMEMORY;
    }
    const auto freeOutputDocument = wil::scope_exit([&] noexcept { yyjson_mut_doc_free(outputDocument); });
    yyjson_mut_val* root = yyjson_mut_obj(outputDocument);
    if (root == nullptr)
    {
        return E_OUTOFMEMORY;
    }
    yyjson_mut_doc_set_root(outputDocument, root);
    const std::string shellUtf8 = Common::Strings::Utf8FromUtf16StrictOrEmpty(next.defaultShell);
    const std::string hyperlinkPolicyUtf8 = Common::Strings::Utf8FromUtf16StrictOrEmpty(next.hyperlinkPolicy);
    const std::string osc52PolicyUtf8 = Common::Strings::Utf8FromUtf16StrictOrEmpty(next.osc52Policy);
    const bool built = yyjson_mut_obj_add_strcpy(outputDocument, root, "defaultShell", shellUtf8.c_str()) &&
        yyjson_mut_obj_add_strcpy(outputDocument, root, "fontFamily", fontUtf8.c_str()) &&
        yyjson_mut_obj_add_real(outputDocument, root, "fontSizeDip", static_cast<double>(next.fontSizeDip)) &&
        yyjson_mut_obj_add_uint(outputDocument, root, "maxFormattedMiB", next.maxFormattedBytes / (1024u * 1024u)) &&
        yyjson_mut_obj_add_uint(outputDocument, root, "pasteMaxBytes", next.pasteMaxBytes) &&
        yyjson_mut_obj_add_strcpy(outputDocument, root, "warnOnUnsafePaste", next.warnOnUnsafePaste ? "1" : "0") &&
        yyjson_mut_obj_add_strcpy(outputDocument, root, "followPathWhenIdle", next.followPathWhenIdle ? "1" : "0") &&
        yyjson_mut_obj_add_strcpy(outputDocument, root, "hyperlinkPolicy", hyperlinkPolicyUtf8.c_str()) &&
        yyjson_mut_obj_add_strcpy(outputDocument, root, "osc52Policy", osc52PolicyUtf8.c_str()) &&
        yyjson_mut_obj_add_uint(outputDocument, root, "osc52MaxBytes", next.osc52MaxBytes);
    if (! built)
    {
        return E_OUTOFMEMORY;
    }
    char* written = yyjson_mut_write(outputDocument, YYJSON_WRITE_NOFLAG, nullptr);
    if (written == nullptr)
    {
        return E_OUTOFMEMORY;
    }
    const auto freeWritten = wil::scope_exit([&] noexcept { free(written); });

    _config = std::move(next);
    if (_windowHandle.load(std::memory_order_acquire) == nullptr)
    {
        _liveFontSizeDip = _config.fontSizeDip;
    }
    _followPathWhenIdle.store(_config.followPathWhenIdle, std::memory_order_release);
    _osc52Policy.store(_config.osc52Policy == L"deny"   ? Osc52Policy::Deny
                           : _config.osc52Policy == L"allow" ? Osc52Policy::Allow
                                                               : Osc52Policy::Ask,
                       std::memory_order_release);
    _osc52MaxBytes.store(_config.osc52MaxBytes, std::memory_order_release);
    {
        std::scoped_lock stateLock(_stateMutex);
        if (! _followPathWhenIdle.load(std::memory_order_acquire))
        {
            _pendingFollowDirectory.clear();
            _followState = TerminalFollowState::Disabled;
        }
        else if (_shellKind == ShellKind::PowerShell &&
                 (_sourceLocationKind == TerminalLocationKind::WindowsLocal ||
                  _sourceLocationKind == TerminalLocationKind::WindowsUnc ||
                  _sourceLocationKind == TerminalLocationKind::PluginBacked) &&
                 IsSupportedAbsoluteWindowsPath(_sourceLocationPath))
        {
            if (_integrationTrusted &&
                OrdinalString::EqualsNoCase(_trustedCurrentDirectory, _sourceLocationPath))
            {
                _pendingFollowDirectory.clear();
                _followState = TerminalFollowState::Applied;
            }
            else
            {
                _pendingFollowDirectory = _sourceLocationPath;
                _followState = TerminalFollowState::Pending;
            }
        }
        else
        {
            _pendingFollowDirectory.clear();
            _followState = _sourceLocationPath.empty() ? TerminalFollowState::Detached
                                                       : TerminalFollowState::Paused;
        }
        _stateGeneration.fetch_add(1u, std::memory_order_acq_rel);
    }
    _configurationJson = written;
    discardDeviceResources();
    resizeTerminal();
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
    return S_OK;
}

HRESULT STDMETHODCALLTYPE Terminal::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }
    *configurationJsonUtf8 = _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE Terminal::SomethingToSave(BOOL* somethingToSave) noexcept
{
    if (somethingToSave == nullptr)
    {
        return E_POINTER;
    }
    const Config defaults;
    *somethingToSave = (_config.fontFamily == defaults.fontFamily && _config.fontSizeDip == defaults.fontSizeDip &&
                        _config.maxFormattedBytes == defaults.maxFormattedBytes && _config.defaultShell == defaults.defaultShell &&
                        _config.pasteMaxBytes == defaults.pasteMaxBytes &&
                        _config.osc52MaxBytes == defaults.osc52MaxBytes &&
                        _config.hyperlinkPolicy == defaults.hyperlinkPolicy &&
                        _config.osc52Policy == defaults.osc52Policy &&
                        _config.followPathWhenIdle == defaults.followPathWhenIdle &&
                        _config.warnOnUnsafePaste == defaults.warnOnUnsafePaste)
        ? FALSE
        : TRUE;
    return S_OK;
}

bool Terminal::validateLocation(const TerminalLogicalLocation& location) noexcept
{
    return ValidateTerminalLogicalLocation(location);
}

std::wstring Terminal::copySpan(TerminalUtf16Span span)
{
    return CopyTerminalSpan(span);
}

HRESULT STDMETHODCALLTYPE Terminal::Open(const TerminalOpenContext* context) noexcept
{
    if (context == nullptr || context->sizeBytes < sizeof(TerminalOpenContext) ||
        context->parentWindow == nullptr || ! IsWindow(context->parentWindow) || ! validateLocation(context->sourceLocation) ||
        ! validateLocation(context->launchLocation))
    {
        return E_INVALIDARG;
    }
    if (_window || _closing.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED);
    }
    if (const HRESULT closeWorkHr = ensureCloseWork(); FAILED(closeWorkHr))
    {
        return closeWorkHr;
    }

    _instanceId = context->instanceId;
    _liveFontSizeDip = _config.fontSizeDip;
    _exitCodePresent.store(false, std::memory_order_release);
    _exitCode.store(0u, std::memory_order_release);
    _finalSnapshotComplete.store(false, std::memory_order_release);
    _rootProcessExited.store(false, std::memory_order_release);
    _rootExitObservedTimestampNs.store(0, std::memory_order_release);
    {
        std::scoped_lock callbackLock(_eventCallbackMutex);
        _eventDeliveredSessionGeneration = 0u;
    }
    {
        std::scoped_lock stateLock(_stateMutex);
        _originalSource = context->originalSource;
        _sourceGeneration = context->sourceGeneration;
        _sourceLocationKind = context->sourceLocation.kind;
        _sourceLocationPath = CopyTerminalLocationPath(context->sourceLocation);
        _sourcePluginShortId = copySpan(context->sourceLocation.pluginShortId);
        _launchWslDistribution = copySpan(context->launchLocation.wslDistribution);
        _trustedCurrentDirectory.clear();
        _trustedHistoryPath.clear();
        SecureZeroMemory(_trustedPromptBuffer.data(), _trustedPromptBuffer.size() * sizeof(wchar_t));
        _trustedPromptBuffer.clear();
        ScrubStrings(_trustedCurrentSessionHistory);
        SecureZeroMemory(_pendingPromptExpectedBuffer.data(), _pendingPromptExpectedBuffer.size() * sizeof(wchar_t));
        SecureZeroMemory(_pendingPromptReplacement.data(), _pendingPromptReplacement.size() * sizeof(wchar_t));
        _pendingPromptExpectedBuffer.clear();
        _pendingPromptReplacement.clear();
        _pendingFollowDirectory.clear();
        _followState = _followPathWhenIdle.load(std::memory_order_acquire) ? TerminalFollowState::Pending
                                                                          : TerminalFollowState::Disabled;
        _integrationTrusted = false;
        _trustedHistoryProviderAvailable = false;
        _idleAtPrimaryPrompt = false;
        _hasPendingUserInput = false;
        _commandSubmitted = false;
    }
    HRESULT hr = createWindow(context->parentWindow);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = initializeEngine();
    if (FAILED(hr))
    {
        const std::wstring& detail = _engineInitializationError.empty() ? _runtime.ErrorText() : _engineInitializationError;
        setDiagnostic(std::format(L"{}\r\n\r\n{}", LoadStringResource(g_hInstance, IDS_TERMINAL_RUNTIME_ERROR), detail));
        setLifecycle(TerminalLifecycleState::Diagnostic);
        return S_OK;
    }

    const TerminalLocationKind locationKind = context->launchLocation.kind;
    const bool supportedWindows = (locationKind == TerminalLocationKind::WindowsLocal || locationKind == TerminalLocationKind::WindowsUnc) &&
        ! copySpan(context->launchLocation.windowsPath).empty();
    const bool supportedWsl = locationKind == TerminalLocationKind::Wsl && ! copySpan(context->launchLocation.wslDistribution).empty() &&
        ! copySpan(context->launchLocation.wslAbsolutePath).empty();
    if (! supportedWindows && ! supportedWsl)
    {
        setDiagnostic(LoadStringResource(g_hInstance, IDS_TERMINAL_UNSUPPORTED_LOCATION));
        setLifecycle(TerminalLifecycleState::Diagnostic);
        return S_OK;
    }

    _sessionGeneration.fetch_add(1u, std::memory_order_acq_rel);
    setLifecycle(TerminalLifecycleState::Running);
    hr = startPseudoConsole(context->launchLocation);
    if (FAILED(hr))
    {
        setDiagnostic(FormatStringResource(g_hInstance, IDS_TERMINAL_LAUNCH_ERROR, static_cast<uint32_t>(hr)));
        setLifecycle(TerminalLifecycleState::Failed);
        return S_OK;
    }

    resizeTerminal();
    SetFocus(_window.get());
    return S_OK;
}

HRESULT Terminal::createWindow(HWND parent) noexcept
{
    {
        std::scoped_lock lock(g_terminalWindowClassMutex);
        WNDCLASSW existing{};
        if (GetClassInfoW(g_hInstance, kTerminalWindowClass, &existing) != FALSE)
        {
            if (existing.lpfnWndProc != &Terminal::WindowProc)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            g_terminalWindowClassRegistered = true;
        }
        else
        {
            WNDCLASSW windowClass{};
            windowClass.style = CS_DBLCLKS;
            windowClass.lpfnWndProc = &Terminal::WindowProc;
            windowClass.hInstance = g_hInstance;
            windowClass.hCursor = LoadCursorW(nullptr, IDC_IBEAM);
            windowClass.lpszClassName = kTerminalWindowClass;
            if (RegisterClassW(&windowClass) == 0u)
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
            g_terminalWindowClassRegistered = true;
        }
    }

    _dpi = GetDpiForWindow(parent);
    _window.reset(CreateWindowExW(0,
                                  kTerminalWindowClass,
                                  nullptr,
                                  WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_TABSTOP | WS_VSCROLL,
                                  0,
                                  0,
                                  0,
                                  0,
                                  parent,
                                  nullptr,
                                  g_hInstance,
                                  this));
    if (! _window)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    _windowHandle.store(_window.get(), std::memory_order_release);
    applyNativeScrollbarTheme(_window.get());
    const SCROLLINFO initialScrollInfo = MakeTerminalScrollInfo(
        GhosttyTerminalScrollbar{.total = _rows, .offset = 0u, .len = _rows});
    SetScrollInfo(_window.get(), SB_VERT, &initialScrollInfo, TRUE);
    _accessibility = std::make_unique<TerminalAccessibility>();
    const HRESULT accessibilityHr = _accessibility->Initialize(_window.get());
    if (FAILED(accessibilityHr))
    {
        _accessibility.reset();
        _windowHandle.store(nullptr, std::memory_order_release);
        _window.reset();
        return accessibilityHr;
    }
    ShowWindow(_window.get(), SW_SHOWNA);
    return S_OK;
}

HRESULT Terminal::initializeEngine() noexcept
{
    const bool alreadyInitialized = _runtime.IsLoaded() || _ghosttyTerminal != nullptr || _formatter != nullptr ||
        _renderState != nullptr || _renderRows != nullptr || _renderCells != nullptr || _keyEncoder != nullptr ||
        _keyEvent != nullptr || _mouseEncoder != nullptr || _mouseEvent != nullptr || _kittyPlacementIterator != nullptr;
    if (alreadyInitialized)
    {
        const HRESULT hr = HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED);
        _engineInitializationError = FormatStringResource(
            g_hInstance, IDS_TERMINAL_RUNTIME_INITIALIZATION_ERROR, 0u, static_cast<uint32_t>(hr));
        return hr;
    }

    _engineInitializationError.clear();
#if defined(ENABLE_TESTS)
    const auto injectFailure = [this](uint32_t stage) noexcept
    {
        return _debugEngineInitializationFailStage.load(std::memory_order_acquire) == stage;
    };
#else
    const auto injectFailure = [](uint32_t stage) noexcept
    {
        static_cast<void>(stage);
        return false;
    };
#endif
    const auto fail = [this](uint32_t stage, HRESULT hr) noexcept
    {
        _engineInitializationError = FormatStringResource(
            g_hInstance, IDS_TERMINAL_RUNTIME_INITIALIZATION_ERROR, stage, static_cast<uint32_t>(hr));
        releaseEngine();
        return hr;
    };

    HRESULT hr = _runtime.Load();
    if (FAILED(hr))
    {
        return hr;
    }
    if (injectFailure(1u))
    {
        return fail(1u, E_FAIL);
    }
    if (_runtime.terminalNew(nullptr, &_ghosttyTerminal, _columns, _rows) != GHOSTTY_SUCCESS ||
        _ghosttyTerminal == nullptr || injectFailure(2u))
    {
        return fail(2u, E_FAIL);
    }

    GhosttyFormatterTerminalOptions options{};
    options.size = sizeof(options);
    options.emit = GHOSTTY_FORMATTER_FORMAT_PLAIN;
    options.unwrap = false;
    options.trim = true;
    options.extra.size = sizeof(options.extra);
    options.extra.screen.size = sizeof(options.extra.screen);
    if (_runtime.formatterTerminalNew(nullptr, &_formatter, _ghosttyTerminal, options) != GHOSTTY_SUCCESS ||
        _formatter == nullptr || injectFailure(3u))
    {
        return fail(3u, E_FAIL);
    }

    if (_runtime.renderStateNew(nullptr, &_renderState) != GHOSTTY_SUCCESS || _renderState == nullptr || injectFailure(4u))
    {
        return fail(4u, E_FAIL);
    }
    if (_runtime.renderRowIteratorNew(nullptr, &_renderRows) != GHOSTTY_SUCCESS || _renderRows == nullptr ||
        injectFailure(5u))
    {
        return fail(5u, E_FAIL);
    }
    if (_runtime.renderRowCellsNew(nullptr, &_renderCells) != GHOSTTY_SUCCESS || _renderCells == nullptr ||
        injectFailure(6u))
    {
        return fail(6u, E_FAIL);
    }
    if (_runtime.keyEncoderNew(nullptr, &_keyEncoder) != GHOSTTY_SUCCESS || _keyEncoder == nullptr || injectFailure(7u))
    {
        return fail(7u, E_FAIL);
    }
    if (_runtime.keyEventNew(nullptr, &_keyEvent) != GHOSTTY_SUCCESS || _keyEvent == nullptr || injectFailure(8u))
    {
        return fail(8u, E_FAIL);
    }
    if (_runtime.mouseEncoderNew(nullptr, &_mouseEncoder) != GHOSTTY_SUCCESS || _mouseEncoder == nullptr ||
        injectFailure(9u))
    {
        return fail(9u, E_FAIL);
    }
    if (_runtime.mouseEventNew(nullptr, &_mouseEvent) != GHOSTTY_SUCCESS || _mouseEvent == nullptr || injectFailure(10u))
    {
        return fail(10u, E_FAIL);
    }
    if (_runtime.kittyPlacementIteratorNew(nullptr, &_kittyPlacementIterator) != GHOSTTY_SUCCESS ||
        _kittyPlacementIterator == nullptr || injectFailure(11u))
    {
        return fail(11u, E_FAIL);
    }

    if (_runtime.terminalSet(_ghosttyTerminal, GHOSTTY_TERMINAL_OPT_USERDATA, this) != GHOSTTY_SUCCESS || injectFailure(12u))
    {
        return fail(12u, E_FAIL);
    }
    if (_runtime.terminalSet(
            _ghosttyTerminal,
            GHOSTTY_TERMINAL_OPT_WRITE_PTY,
            reinterpret_cast<const void*>(&Terminal::GhosttyWritePty)) != GHOSTTY_SUCCESS || injectFailure(13u))
    {
        return fail(13u, E_FAIL);
    }
    if (_runtime.terminalSet(_ghosttyTerminal,
                             GHOSTTY_TERMINAL_OPT_CLIPBOARD_WRITE,
                             reinterpret_cast<const void*>(&Terminal::GhosttyClipboardWrite)) != GHOSTTY_SUCCESS ||
        injectFailure(14u))
    {
        return fail(14u, E_FAIL);
    }
    const size_t clipboardWriteMaxBytes = kGhosttyClipboardWriteMaxBytes;
    const bool kittyClipboardWriteEnabled = false;
    if (_runtime.terminalSet(_ghosttyTerminal, GHOSTTY_TERMINAL_OPT_CLIPBOARD_READ, nullptr) != GHOSTTY_SUCCESS ||
        injectFailure(15u))
    {
        return fail(15u, E_FAIL);
    }
    if (_runtime.terminalSet(
            _ghosttyTerminal, GHOSTTY_TERMINAL_OPT_CLIPBOARD_WRITE_MAX_BYTES, &clipboardWriteMaxBytes) != GHOSTTY_SUCCESS ||
        injectFailure(16u))
    {
        return fail(16u, E_FAIL);
    }
    if (_runtime.terminalSet(
            _ghosttyTerminal, GHOSTTY_TERMINAL_OPT_KITTY_CLIPBOARD_WRITE_ENABLED, &kittyClipboardWriteEnabled) != GHOSTTY_SUCCESS ||
        injectFailure(17u))
    {
        return fail(17u, E_FAIL);
    }
    const uint64_t storageLimit = kKittyImageStorageBytes;
    const size_t apcLimit = kMaximumInputBytes;
    const bool externalMedium = false;
    const GhosttyKittyPlacementLimits placementLimits{
        sizeof(GhosttyKittyPlacementLimits), kMaximumKittyPlacements, kMaximumKittyPlacementBytes};
    if (_runtime.terminalSet(_ghosttyTerminal, GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_STORAGE_LIMIT, &storageLimit) != GHOSTTY_SUCCESS ||
        injectFailure(18u))
    {
        return fail(18u, E_FAIL);
    }
    if (_runtime.terminalSet(_ghosttyTerminal, GHOSTTY_TERMINAL_OPT_KITTY_PLACEMENT_LIMITS, &placementLimits) != GHOSTTY_SUCCESS ||
        injectFailure(19u))
    {
        return fail(19u, E_FAIL);
    }
    if (_runtime.terminalSet(_ghosttyTerminal, GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_FILE, &externalMedium) != GHOSTTY_SUCCESS ||
        injectFailure(20u))
    {
        return fail(20u, E_FAIL);
    }
    if (_runtime.terminalSet(_ghosttyTerminal, GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_TEMP_FILE, nullptr) != GHOSTTY_SUCCESS ||
        injectFailure(21u))
    {
        return fail(21u, E_FAIL);
    }
    if (_runtime.terminalSet(_ghosttyTerminal, GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_SHARED_MEM, &externalMedium) != GHOSTTY_SUCCESS ||
        injectFailure(22u))
    {
        return fail(22u, E_FAIL);
    }
    if (_runtime.terminalSet(_ghosttyTerminal, GHOSTTY_TERMINAL_OPT_APC_MAX_BYTES_KITTY, &apcLimit) != GHOSTTY_SUCCESS ||
        injectFailure(23u))
    {
        return fail(23u, E_FAIL);
    }
    hr = _kittyPipeline.Start(&Terminal::NotifyKittyReady, this);
    if (FAILED(hr) || injectFailure(24u))
    {
        return fail(24u, FAILED(hr) ? hr : E_FAIL);
    }
    return S_OK;
}

void Terminal::releaseEngine() noexcept
{
    _kittyPipeline.Stop();
    {
        std::scoped_lock lock(_terminalMutex);
        if (_kittyPlacementIterator != nullptr && _runtime.kittyPlacementIteratorFree != nullptr)
        {
            _runtime.kittyPlacementIteratorFree(_kittyPlacementIterator);
            _kittyPlacementIterator = nullptr;
        }
        if (_mouseEvent != nullptr && _runtime.mouseEventFree != nullptr)
        {
            _runtime.mouseEventFree(_mouseEvent);
            _mouseEvent = nullptr;
        }
        if (_mouseEncoder != nullptr && _runtime.mouseEncoderFree != nullptr)
        {
            _runtime.mouseEncoderFree(_mouseEncoder);
            _mouseEncoder = nullptr;
        }
        if (_keyEvent != nullptr && _runtime.keyEventFree != nullptr)
        {
            _runtime.keyEventFree(_keyEvent);
            _keyEvent = nullptr;
        }
        if (_keyEncoder != nullptr && _runtime.keyEncoderFree != nullptr)
        {
            _runtime.keyEncoderFree(_keyEncoder);
            _keyEncoder = nullptr;
        }
        if (_renderCells != nullptr && _runtime.renderRowCellsFree != nullptr)
        {
            _runtime.renderRowCellsFree(_renderCells);
            _renderCells = nullptr;
        }
        if (_renderRows != nullptr && _runtime.renderRowIteratorFree != nullptr)
        {
            _runtime.renderRowIteratorFree(_renderRows);
            _renderRows = nullptr;
        }
        if (_renderState != nullptr && _runtime.renderStateFree != nullptr)
        {
            _runtime.renderStateFree(_renderState);
            _renderState = nullptr;
        }
        if (_formatter != nullptr && _runtime.formatterFree != nullptr)
        {
            _runtime.formatterFree(_formatter);
            _formatter = nullptr;
        }
        if (_ghosttyTerminal != nullptr && _runtime.terminalFree != nullptr)
        {
            _runtime.terminalFree(_ghosttyTerminal);
            _ghosttyTerminal = nullptr;
        }
    }
    _runtime.Unload();
}
























void Terminal::drawCommandSurface(float widthDip, float heightDip) noexcept
{
    if (_commandSurfaceKind == CommandSurfaceKind::None || !_renderTarget || !_foregroundBrush ||
        !_overlayTextFormat || !_overlayBoldTextFormat)
    {
        return;
    }

    const bool find = _commandSurfaceKind == CommandSurfaceKind::Find;
    const float panelLeft = 16.0f;
    const float panelRight = std::max(panelLeft + 240.0f, widthDip - 16.0f);
    const float panelTop = find ? 16.0f : std::max(16.0f, heightDip - 340.0f);
    const float panelBottom = find ? std::min(heightDip - 16.0f, 150.0f) : heightDip - 16.0f;
    const D2D1_RECT_F panel = D2D1::RectF(panelLeft, panelTop, panelRight, panelBottom);

    D2D1_COLOR_F background = RedSalamander::DxUi::ColorFromArgb(_backgroundArgb);
    background.a = _themeHighContrast ? 1.0f : 0.97f;
    _foregroundBrush->SetColor(background);
    _renderTarget->FillRectangle(panel, _foregroundBrush.get());
    _foregroundBrush->SetColor(RedSalamander::DxUi::ColorFromArgb(_hyperlinkArgb));
    _renderTarget->DrawRectangle(panel, _foregroundBrush.get(), 1.0f);

    const std::wstring title = LoadStringResource(g_hInstance, find ? IDS_TERMINAL_FIND_TITLE : IDS_TERMINAL_SUGGESTIONS_TITLE);
    const std::wstring icon = find ? L"⌕" : L"↶";
    const std::wstring heading = std::format(L"{}  {}", icon, title);
    _foregroundBrush->SetColor(RedSalamander::DxUi::ColorFromArgb(_foregroundArgb));
    _renderTarget->DrawTextW(heading.data(), static_cast<UINT32>(heading.size()), _overlayBoldTextFormat.get(),
                             D2D1::RectF(panel.left + 12.0f, panel.top + 8.0f, panel.right - 12.0f, panel.top + 34.0f),
                             _foregroundBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

    const D2D1_RECT_F queryRect = D2D1::RectF(panel.left + 12.0f, panel.top + 36.0f, panel.right - 12.0f, panel.top + 66.0f);
    _foregroundBrush->SetColor(RedSalamander::DxUi::ColorFromArgb(_selectionBackgroundArgb));
    _renderTarget->DrawRectangle(queryRect, _foregroundBrush.get(), 1.0f);
    const std::wstring queryText = _commandSurfaceQuery.empty()
        ? LoadStringResource(g_hInstance, find ? IDS_TERMINAL_FIND_PLACEHOLDER : IDS_TERMINAL_SUGGESTIONS_PLACEHOLDER)
        : _commandSurfaceQuery;
    uint32_t queryColor = _commandSurfaceQuery.empty() ? WithAlpha(_foregroundArgb, 0xA0u) : _foregroundArgb;
    _foregroundBrush->SetColor(RedSalamander::DxUi::ColorFromArgb(queryColor));
    _renderTarget->DrawTextW(queryText.data(), static_cast<UINT32>(queryText.size()), _overlayTextFormat.get(),
                             D2D1::RectF(queryRect.left + 8.0f, queryRect.top + 4.0f, queryRect.right - 8.0f, queryRect.bottom - 2.0f),
                             _foregroundBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

    const float rowsTop = panel.top + 70.0f;
    const float footerHeight = find ? 22.0f : 24.0f;
    const size_t visibleRows = static_cast<size_t>(std::max(0.0f, std::floor((panel.bottom - rowsTop - footerHeight) / 30.0f)));
    const size_t count = find ? _commandSurfaceFindRows.size() : _commandSurfaceSuggestionRows.size();
    const size_t firstRow = visibleRows != 0u && _commandSurfaceSelected >= visibleRows
        ? _commandSurfaceSelected - visibleRows + 1u
        : 0u;
    const size_t endRow = std::min(count, firstRow + visibleRows);
    for (size_t index = firstRow; index < endRow; ++index)
    {
        const float top = rowsTop + static_cast<float>(index - firstRow) * 30.0f;
        const D2D1_RECT_F rowRect = D2D1::RectF(panel.left + 8.0f, top, panel.right - 8.0f, top + 28.0f);
        if (index == _commandSurfaceSelected)
        {
            _foregroundBrush->SetColor(RedSalamander::DxUi::ColorFromArgb(_selectionBackgroundArgb));
            _renderTarget->FillRectangle(rowRect, _foregroundBrush.get());
        }
        std::wstring rowText;
        if (find)
        {
            rowText = std::format(L"{}  {}", _commandSurfaceFindRows[index].lineNumber + 1u, _commandSurfaceFindRows[index].text);
        }
        else
        {
            const auto& suggestion = _commandSurfaceSuggestionRows[index];
            rowText = std::format(L"{}    · {}", suggestion.text,
                                  LoadStringResource(g_hInstance, suggestion.currentSession
                                      ? IDS_TERMINAL_HISTORY_CURRENT_SESSION
                                      : IDS_TERMINAL_HISTORY_PERSISTED));
        }
        _foregroundBrush->SetColor(RedSalamander::DxUi::ColorFromArgb(
            index == _commandSurfaceSelected ? _selectionForegroundArgb : _foregroundArgb));
        _renderTarget->DrawTextW(rowText.data(), static_cast<UINT32>(std::min<size_t>(rowText.size(), UINT32_MAX)),
                                 _overlayTextFormat.get(),
                                 D2D1::RectF(rowRect.left + 6.0f, rowRect.top + 4.0f, rowRect.right - 6.0f, rowRect.bottom),
                                 _foregroundBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
        if (find && !_commandSurfaceQuery.empty())
        {
            const float matchLeft = rowRect.left + 6.0f +
                static_cast<float>(std::to_wstring(_commandSurfaceFindRows[index].lineNumber + 1u).size() + 2u +
                                   _commandSurfaceFindRows[index].matchOffset) * _cellWidthDip;
            const float matchRight = std::min(rowRect.right - 6.0f,
                matchLeft + static_cast<float>(_commandSurfaceQuery.size()) * _cellWidthDip);
            _foregroundBrush->SetColor(RedSalamander::DxUi::ColorFromArgb(_hyperlinkArgb));
            _renderTarget->DrawLine(D2D1::Point2F(matchLeft, rowRect.bottom - 2.0f),
                                    D2D1::Point2F(matchRight, rowRect.bottom - 2.0f), _foregroundBrush.get(), 2.0f);
        }
    }

    std::wstring status;
    if (! _commandSurfaceStatusOverride.empty())
    {
        status = _commandSurfaceStatusOverride;
    }
    else if (_commandSurfaceLoading)
    {
        status = LoadStringResource(g_hInstance, IDS_TERMINAL_SURFACE_LOADING);
    }
    else if (find && !_commandSurfaceQuery.empty() && _commandSurfaceFindRows.empty())
    {
        status = LoadStringResource(g_hInstance, IDS_TERMINAL_SURFACE_NO_MATCHES);
    }
    else if (find && !_commandSurfaceFindRows.empty())
    {
        status = FormatStringResource(g_hInstance,
                                      IDS_TERMINAL_FIND_MATCH_COUNT,
                                      _commandSurfaceSelected + 1u,
                                      _commandSurfaceFindTotal);
    }
    else if (! find && _commandSurfaceSuggestionRows.empty())
    {
        status = ! _commandSurfaceHistoryTrusted ||
                (((! _persistedHistory || _persistedHistory->empty()) &&
                  (! _commandSurfaceSessionHistorySnapshot || _commandSurfaceSessionHistorySnapshot->empty())))
            ? LoadStringResource(g_hInstance, IDS_TERMINAL_HISTORY_UNAVAILABLE)
            : LoadStringResource(g_hInstance, IDS_TERMINAL_SURFACE_NO_MATCHES);
    }
    const std::wstring instructions = LoadStringResource(
        g_hInstance, find ? IDS_TERMINAL_FIND_INSTRUCTIONS : IDS_TERMINAL_SUGGESTIONS_INSTRUCTIONS);
    if (! status.empty())
    {
        status.append(L"    ");
    }
    status.append(instructions);
    _foregroundBrush->SetColor(RedSalamander::DxUi::ColorFromArgb(WithAlpha(_foregroundArgb, 0xC0u)));
    _renderTarget->DrawTextW(status.data(), static_cast<UINT32>(status.size()), _overlayTextFormat.get(),
                             D2D1::RectF(panel.left + 12.0f, panel.bottom - footerHeight, panel.right - 12.0f, panel.bottom - 2.0f),
                             _foregroundBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
    _foregroundBrush->SetColor(RedSalamander::DxUi::ColorFromArgb(_foregroundArgb));
}

std::wstring Terminal::commandSurfaceAccessibilityText() const
{
    if (_commandSurfaceKind == CommandSurfaceKind::None)
    {
        return {};
    }
    const bool find = _commandSurfaceKind == CommandSurfaceKind::Find;
    const std::wstring title = LoadStringResource(g_hInstance, find ? IDS_TERMINAL_FIND_TITLE : IDS_TERMINAL_SUGGESTIONS_TITLE);
    std::wstring state;
    if (! _commandSurfaceStatusOverride.empty())
    {
        state = _commandSurfaceStatusOverride;
    }
    else if (_commandSurfaceLoading)
    {
        state = LoadStringResource(g_hInstance, IDS_TERMINAL_SURFACE_LOADING);
    }
    else if (find && !_commandSurfaceFindRows.empty())
    {
        state = FormatStringResource(g_hInstance, IDS_TERMINAL_FIND_MATCH_COUNT,
                                     _commandSurfaceSelected + 1u, _commandSurfaceFindTotal);
        state.append(L". ");
        state.append(_commandSurfaceFindRows[_commandSurfaceSelected].text);
    }
    else if (! find && !_commandSurfaceSuggestionRows.empty())
    {
        state = _commandSurfaceSuggestionRows[_commandSurfaceSelected].text;
    }
    else
    {
        state = LoadStringResource(g_hInstance, IDS_TERMINAL_SURFACE_NO_MATCHES);
    }
    return FormatStringResource(g_hInstance, IDS_TERMINAL_SURFACE_ACCESSIBILITY,
                                title, _commandSurfaceQuery, state);
}














bool Terminal::openCommandSurface(CommandSurfaceKind kind) noexcept
{
    if (kind == CommandSurfaceKind::None || _commandSurfaceInsertionPending || confirmationVisible() ||
        _closing.load(std::memory_order_acquire))
    {
        return false;
    }
    _commandSurfaceWorker.request_stop();
    retireCommandSurfaceWorker();
    ++_commandSurfaceGeneration;
    _commandSurfaceKind = kind;
    _commandSurfaceLoading = false;
    _commandSurfaceInsertionPending = false;
    _commandSurfaceHistoryTrusted = false;
    _commandSurfaceSessionGeneration = _sessionGeneration.load(std::memory_order_acquire);
    _commandSurfaceStateGeneration = _stateGeneration.load(std::memory_order_acquire);
    SecureZeroMemory(_commandSurfaceStatusOverride.data(), _commandSurfaceStatusOverride.size() * sizeof(wchar_t));
    _commandSurfaceStatusOverride.clear();
    SecureZeroMemory(_commandSurfacePromptSnapshot.data(), _commandSurfacePromptSnapshot.size() * sizeof(wchar_t));
    _commandSurfacePromptSnapshot.clear();
    _commandSurfaceHistoryPath.clear();
    SecureZeroMemory(_commandSurfaceQuery.data(), _commandSurfaceQuery.size() * sizeof(wchar_t));
    _commandSurfaceQuery.clear();
    ScrubFindRows(_commandSurfaceFindRows);
    _commandSurfaceFindTotal = 0u;
    ScrubSuggestionRows(_commandSurfaceSuggestionRows);
    _commandSurfaceSelected = 0u;
    if (_commandSurfaceSessionHistorySnapshot)
    {
        ScrubStrings(*_commandSurfaceSessionHistorySnapshot);
    }
    _commandSurfaceSessionHistorySnapshot.reset();
    if (_commandSurfaceSnapshot)
    {
        SecureZeroMemory(_commandSurfaceSnapshot->data(), _commandSurfaceSnapshot->size() * sizeof(wchar_t));
    }
    if (kind == CommandSurfaceKind::Find)
    {
        _commandSurfaceSnapshot.reset();
        _commandSurfaceSnapshotReady = false;
        const uint64_t snapshotGeneration = ++_commandSurfaceSnapshotGeneration;
        const uint64_t sessionGeneration = _sessionGeneration.load(std::memory_order_acquire);
        const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
        _commandSurfaceLoading = true;
        _commandSurfaceWorker = std::jthread(
            [this, snapshotGeneration, sessionGeneration, hwnd](std::stop_token stopToken) noexcept
            {
                const auto startedAt = std::chrono::steady_clock::now();
                auto snapshot = std::make_shared<std::wstring>(formatScreen());
                if (stopToken.stop_requested())
                {
                    SecureZeroMemory(snapshot->data(), snapshot->size() * sizeof(wchar_t));
                    return;
                }
                auto payload = std::make_unique<TerminalCommandSurfaceReadyPayload>();
                payload->sessionGeneration = sessionGeneration;
                payload->snapshotGeneration = snapshotGeneration;
                payload->snapshotOnly = true;
                payload->snapshot = std::move(snapshot);
                payload->snapshotDurationUs = Debug::Perf::ElapsedUs(startedAt);
                if (hwnd != nullptr)
                {
                    static_cast<void>(PostMessagePayload(
                        hwnd, WndMsg::kTerminalCommandSurfaceReady, 0u, std::move(payload)));
                }
            });
    }
    else
    {
        std::vector<std::wstring> sessionHistory;
        std::wstring promptBuffer;
        std::wstring historyPath;
        {
            std::scoped_lock stateLock(_stateMutex);
            _commandSurfaceStateGeneration = _stateGeneration.load(std::memory_order_acquire);
            _commandSurfaceHistoryTrusted = _integrationTrusted && _trustedHistoryProviderAvailable &&
                _idleAtPrimaryPrompt && _shellKind == ShellKind::PowerShell;
            if (_commandSurfaceHistoryTrusted)
            {
                promptBuffer = _trustedPromptBuffer;
                historyPath = _trustedHistoryPath;
                sessionHistory = _trustedCurrentSessionHistory;
            }
        }
        _commandSurfacePromptSnapshot = std::move(promptBuffer);
        _commandSurfaceHistoryPath = std::move(historyPath);
        _commandSurfaceSessionHistorySnapshot =
            std::make_shared<std::vector<std::wstring>>(std::move(sessionHistory));
        if (_persistedHistory)
        {
            ScrubStrings(*_persistedHistory);
        }
        _persistedHistory = std::make_shared<std::vector<std::wstring>>();
        _persistedHistoryPath = _commandSurfaceHistoryPath;
        if (TerminalCommandSurfaceModel::IsSafeSuggestion(_commandSurfacePromptSnapshot) &&
            _commandSurfacePromptSnapshot.size() <= kMaximumCommandSurfaceQueryCharacters)
        {
            _commandSurfaceQuery = _commandSurfacePromptSnapshot;
        }
        _commandSurfaceSnapshot.reset();
        _commandSurfaceSnapshotReady = false;
    }
    restartCommandSurfaceQuery();
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        publishAccessibilitySnapshot();
        InvalidateRect(hwnd, nullptr, FALSE);
    }
    return true;
}

void Terminal::closeCommandSurface() noexcept
{
    if (_commandSurfaceKind == CommandSurfaceKind::None)
    {
        return;
    }
    ++_commandSurfaceGeneration;
    ++_commandSurfaceSnapshotGeneration;
    retireCommandSurfaceWorker();
    _commandSurfaceKind = CommandSurfaceKind::None;
    _commandSurfaceLoading = false;
    _commandSurfaceInsertionPending = false;
    _commandSurfaceHistoryTrusted = false;
    SecureZeroMemory(_commandSurfaceStatusOverride.data(), _commandSurfaceStatusOverride.size() * sizeof(wchar_t));
    _commandSurfaceStatusOverride.clear();
    SecureZeroMemory(_commandSurfacePromptSnapshot.data(), _commandSurfacePromptSnapshot.size() * sizeof(wchar_t));
    _commandSurfacePromptSnapshot.clear();
    _commandSurfaceHistoryPath.clear();
    {
        std::scoped_lock stateLock(_stateMutex);
        const bool hadPendingReplacement = ! _pendingPromptReplacement.empty();
        SecureZeroMemory(_pendingPromptExpectedBuffer.data(), _pendingPromptExpectedBuffer.size() * sizeof(wchar_t));
        SecureZeroMemory(_pendingPromptReplacement.data(), _pendingPromptReplacement.size() * sizeof(wchar_t));
        _pendingPromptExpectedBuffer.clear();
        _pendingPromptReplacement.clear();
        if (hadPendingReplacement)
        {
            _stateGeneration.fetch_add(1u, std::memory_order_acq_rel);
        }
    }
    SecureZeroMemory(_commandSurfaceQuery.data(), _commandSurfaceQuery.size() * sizeof(wchar_t));
    _commandSurfaceQuery.clear();
    if (_commandSurfaceSnapshot)
    {
        SecureZeroMemory(_commandSurfaceSnapshot->data(), _commandSurfaceSnapshot->size() * sizeof(wchar_t));
    }
    _commandSurfaceSnapshot.reset();
    if (_commandSurfaceSessionHistorySnapshot)
    {
        ScrubStrings(*_commandSurfaceSessionHistorySnapshot);
    }
    _commandSurfaceSessionHistorySnapshot.reset();
    _commandSurfaceSnapshotReady = false;
    ScrubFindRows(_commandSurfaceFindRows);
    _commandSurfaceFindTotal = 0u;
    ScrubSuggestionRows(_commandSurfaceSuggestionRows);
    _commandSurfaceSelected = 0u;
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        publishAccessibilitySnapshot();
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void Terminal::restartCommandSurfaceQuery() noexcept
{
    if (_commandSurfaceKind == CommandSurfaceKind::None)
    {
        return;
    }
    if (_commandSurfaceKind == CommandSurfaceKind::Find && ! _commandSurfaceSnapshotReady)
    {
        _commandSurfaceSelected = 0u;
        _commandSurfaceLoading = true;
        return;
    }

    retireCommandSurfaceWorker();
    const uint64_t generation = ++_commandSurfaceGeneration;
    const uint64_t sessionGeneration = _sessionGeneration.load(std::memory_order_acquire);
    const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
    const bool suggestions = _commandSurfaceKind == CommandSurfaceKind::Suggestions;
    const std::wstring query = _commandSurfaceQuery;
    const std::shared_ptr<std::wstring> snapshot = suggestions ? nullptr : _commandSurfaceSnapshot;
    const std::shared_ptr<std::vector<std::wstring>> persistedHistory = _persistedHistory;
    const std::shared_ptr<std::vector<std::wstring>> sessionHistory = _commandSurfaceSessionHistorySnapshot;
    const std::wstring historyPath = _commandSurfaceHistoryPath;
    const bool loadPowerShellHistory = suggestions &&
        _commandSurfaceHistoryTrusted && ! historyPath.empty() &&
        (!persistedHistory || persistedHistory->empty());

    _commandSurfaceSelected = 0u;
    _commandSurfaceLoading = suggestions || ! query.empty();
    if (! suggestions && query.empty())
    {
        _commandSurfaceFindRows.clear();
        _commandSurfaceFindTotal = 0u;
        _commandSurfaceLoading = false;
        return;
    }

    const auto startedAt = std::chrono::steady_clock::now();
    _commandSurfaceWorker = std::jthread(
        [generation,
         sessionGeneration,
         hwnd,
         suggestions,
         query = std::wstring(query),
         snapshot,
         persistedHistory,
         sessionHistory,
         historyPath,
         loadPowerShellHistory,
         startedAt](std::stop_token stopToken) mutable noexcept
        {
            const auto scrubQuery = wil::scope_exit([&query]() noexcept
            {
                SecureZeroMemory(query.data(), query.size() * sizeof(wchar_t));
            });
            auto payload = std::make_unique<TerminalCommandSurfaceReadyPayload>();
            payload->generation = generation;
            payload->sessionGeneration = sessionGeneration;
            payload->suggestions = suggestions;
            if (suggestions)
            {
                std::shared_ptr<std::vector<std::wstring>> effectivePersistedHistory = persistedHistory;
                if (loadPowerShellHistory)
                {
                    const auto historyLoadStartedAt = std::chrono::steady_clock::now();
                    effectivePersistedHistory = std::make_shared<std::vector<std::wstring>>(
                        TerminalCommandSurfaceModel::LoadPowerShellHistory(
                            historyPath,
                            kMaximumPersistedHistoryBytes,
                            kMaximumPersistedHistoryEntries,
                            stopToken));
                    payload->historyLoadDurationUs = Debug::Perf::ElapsedUs(historyLoadStartedAt);
                }
                if (stopToken.stop_requested())
                {
                    return;
                }
                payload->persistedHistory = effectivePersistedHistory;
                static const std::vector<std::wstring> emptyHistory;
                payload->suggestion = TerminalCommandSurfaceModel::FilterSuggestions(
                    effectivePersistedHistory ? *effectivePersistedHistory : emptyHistory,
                    sessionHistory ? *sessionHistory : emptyHistory,
                    query,
                    kMaximumSuggestionRows,
                    stopToken);
            }
            else
            {
                payload->find = TerminalCommandSurfaceModel::Find(
                    snapshot ? std::wstring_view(*snapshot) : std::wstring_view{},
                    query,
                    kMaximumFindResultRows,
                    stopToken);
            }
            if (stopToken.stop_requested() || (suggestions ? payload->suggestion.cancelled : payload->find.cancelled))
            {
                return;
            }
            payload->durationUs = Debug::Perf::ElapsedUs(startedAt);
            if (hwnd != nullptr)
            {
                static_cast<void>(PostMessagePayload(
                    hwnd, WndMsg::kTerminalCommandSurfaceReady, 0u, std::move(payload)));
            }
        });
}

void Terminal::navigateCommandSurface(int direction) noexcept
{
    const size_t count = _commandSurfaceKind == CommandSurfaceKind::Find
        ? _commandSurfaceFindRows.size()
        : _commandSurfaceSuggestionRows.size();
    if (count == 0u)
    {
        return;
    }
    if (direction >= 0)
    {
        _commandSurfaceSelected = (_commandSurfaceSelected + 1u) % count;
    }
    else
    {
        _commandSurfaceSelected = (_commandSurfaceSelected + count - 1u) % count;
    }
    if (_commandSurfaceKind == CommandSurfaceKind::Find)
    {
        scrollViewportTo(_commandSurfaceFindRows[_commandSurfaceSelected].lineNumber);
    }
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        publishAccessibilitySnapshot();
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void Terminal::activateCommandSurfaceSelection() noexcept
{
    if (_commandSurfaceKind == CommandSurfaceKind::Find)
    {
        if (!_commandSurfaceFindRows.empty())
        {
            scrollViewportTo(_commandSurfaceFindRows[_commandSurfaceSelected].lineNumber);
            navigateCommandSurface(1);
        }
        return;
    }
    if (_commandSurfaceKind != CommandSurfaceKind::Suggestions ||
        _commandSurfaceSelected >= _commandSurfaceSuggestionRows.size())
    {
        return;
    }
    const std::wstring& insertion = _commandSurfaceSuggestionRows[_commandSurfaceSelected].text;
    const std::string insertionUtf8 = Common::Strings::Utf8FromUtf16ReplacingInvalid(insertion);
    bool accepted = ! insertionUtf8.empty() && insertionUtf8.size() <= _config.pasteMaxBytes &&
        TerminalCommandSurfaceModel::IsSafeSuggestion(insertion) &&
        _commandSurfaceSessionGeneration == _sessionGeneration.load(std::memory_order_acquire) &&
        _commandSurfaceStateGeneration == _stateGeneration.load(std::memory_order_acquire);
    {
        std::scoped_lock stateLock(_stateMutex);
        accepted = accepted && _integrationTrusted && _trustedHistoryProviderAvailable && _idleAtPrimaryPrompt &&
            _commandSurfaceStateGeneration == _stateGeneration.load(std::memory_order_acquire) &&
            _trustedPromptBuffer == _commandSurfacePromptSnapshot && _pendingPromptReplacement.empty();
        if (accepted)
        {
            _pendingPromptExpectedBuffer = _trustedPromptBuffer;
            _pendingPromptReplacement = insertion;
            _stateGeneration.fetch_add(1u, std::memory_order_acq_rel);
        }
    }
    if (accepted)
    {
        _commandSurfaceInsertionPending = true;
        _commandSurfaceLoading = true;
        _commandSurfaceStatusOverride.clear();
    }
    else
    {
        _commandSurfaceInsertionPending = false;
        _commandSurfaceLoading = false;
        _commandSurfaceStatusOverride = LoadStringResource(g_hInstance, IDS_TERMINAL_HISTORY_PROMPT_CHANGED);
    }
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        publishAccessibilitySnapshot();
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

bool Terminal::handleCommandSurfaceKey(WPARAM virtualKey) noexcept
{
    if (_commandSurfaceKind == CommandSurfaceKind::None)
    {
        return false;
    }
    if (_commandSurfaceInsertionPending)
    {
        return true;
    }
    switch (virtualKey)
    {
        case VK_ESCAPE:
            MarkTranslatedCharacterConsumed(_translatedCharacterSuppressionPending);
            closeCommandSurface();
            break;
        case VK_UP: navigateCommandSurface(-1); break;
        case VK_DOWN: navigateCommandSurface(1); break;
        case VK_PRIOR:
            for (size_t index = 0u; index < 5u; ++index) navigateCommandSurface(-1);
            break;
        case VK_NEXT:
            for (size_t index = 0u; index < 5u; ++index) navigateCommandSurface(1);
            break;
        case VK_RETURN:
            MarkTranslatedCharacterConsumed(_translatedCharacterSuppressionPending);
            activateCommandSurfaceSelection();
            break;
        case VK_TAB:
            MarkTranslatedCharacterConsumed(_translatedCharacterSuppressionPending);
            activateCommandSurfaceSelection();
            break;
        case VK_F3:
            navigateCommandSurface((GetKeyState(VK_SHIFT) & 0x8000) != 0 ? -1 : 1);
            break;
        default: break;
    }
    return true;
}

bool Terminal::handleCommandSurfaceCharacter(wchar_t character) noexcept
{
    if (_commandSurfaceKind == CommandSurfaceKind::None)
    {
        return false;
    }
    if (_commandSurfaceInsertionPending)
    {
        return true;
    }
    SecureZeroMemory(_commandSurfaceStatusOverride.data(), _commandSurfaceStatusOverride.size() * sizeof(wchar_t));
    _commandSurfaceStatusOverride.clear();
    if (character == L'\b')
    {
        if (!_commandSurfaceQuery.empty())
        {
            _commandSurfaceQuery.pop_back();
            restartCommandSurfaceQuery();
        }
    }
    else if (character >= L' ' && character != L'\x7F' &&
             _commandSurfaceQuery.size() < kMaximumCommandSurfaceQueryCharacters)
    {
        _commandSurfaceQuery.push_back(character);
        restartCommandSurfaceQuery();
    }
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
    return true;
}

bool Terminal::handleCommandSurfaceClick(POINT clientPoint, bool activate) noexcept
{
    if (_commandSurfaceKind == CommandSurfaceKind::None)
    {
        return false;
    }
    RECT client{};
    const HWND hwnd = _windowHandle.load(std::memory_order_acquire);
    if (hwnd == nullptr || GetClientRect(hwnd, &client) == FALSE)
    {
        return true;
    }
    const float scale = 96.0f / static_cast<float>(std::max<UINT>(_dpi, 1u));
    const float widthDip = static_cast<float>(client.right) * scale;
    const float heightDip = static_cast<float>(client.bottom) * scale;
    const float panelLeft = 16.0f;
    const float panelRight = std::max(panelLeft + 240.0f, widthDip - 16.0f);
    const float panelTop = _commandSurfaceKind == CommandSurfaceKind::Find ? 16.0f : std::max(16.0f, heightDip - 340.0f);
    const float panelBottom = _commandSurfaceKind == CommandSurfaceKind::Find ? std::min(heightDip - 16.0f, 150.0f) : heightDip - 16.0f;
    const D2D1_POINT_2F point{static_cast<float>(clientPoint.x) * scale, static_cast<float>(clientPoint.y) * scale};
    if (point.x < panelLeft || point.x >= panelRight || point.y < panelTop || point.y >= panelBottom)
    {
        closeCommandSurface();
        return true;
    }
    const float rowsTop = panelTop + 70.0f;
    if (point.y >= rowsTop)
    {
        const size_t count = _commandSurfaceKind == CommandSurfaceKind::Find
            ? _commandSurfaceFindRows.size()
            : _commandSurfaceSuggestionRows.size();
        const float footerHeight = _commandSurfaceKind == CommandSurfaceKind::Find ? 22.0f : 24.0f;
        const size_t visibleRows = static_cast<size_t>(
            std::max(0.0f, std::floor((panelBottom - rowsTop - footerHeight) / 30.0f)));
        const size_t firstRow = visibleRows != 0u && _commandSurfaceSelected >= visibleRows
            ? _commandSurfaceSelected - visibleRows + 1u
            : 0u;
        const size_t index = firstRow + static_cast<size_t>((point.y - rowsTop) / 30.0f);
        if (index < count)
        {
            _commandSurfaceSelected = index;
            if (activate)
            {
                activateCommandSurfaceSelection();
            }
            else
            {
                publishAccessibilitySnapshot();
                InvalidateRect(hwnd, nullptr, FALSE);
            }
        }
    }
    return true;
}

































void Terminal::setDiagnostic(std::wstring text) noexcept
{
    {
        std::scoped_lock lock(_terminalMutex);
        _diagnosticText = std::move(text);
    }
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void Terminal::setLifecycle(TerminalLifecycleState lifecycle) noexcept
{
    _lifecycle.store(lifecycle, std::memory_order_release);
    _stateGeneration.fetch_add(1u, std::memory_order_acq_rel);
}

void Terminal::applyNativeScrollbarTheme(HWND hwnd) const noexcept
{
    // High Contrast must clear the Explorer visual style; the shared light/dark
    // helper does not have that terminal-specific policy distinction.
    const wchar_t* themeName = _themeHighContrast ? L"" : (_themeDark ? L"DarkMode_Explorer" : L"Explorer");
    static_cast<void>(SetWindowTheme(hwnd, themeName, nullptr));
    SendMessageW(hwnd, WM_THEMECHANGED, 0, 0);
    RedrawWindow(hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_FRAME);
}

HRESULT STDMETHODCALLTYPE Terminal::GetChildWindow(HWND* childWindow) noexcept
{
    if (childWindow == nullptr)
    {
        return E_POINTER;
    }
    *childWindow = _windowHandle.load(std::memory_order_acquire);
    return *childWindow != nullptr ? S_OK : S_FALSE;
}

HRESULT STDMETHODCALLTYPE Terminal::SetTheme(const TerminalTheme* theme) noexcept
{
    if (theme == nullptr || theme->sizeBytes < sizeof(TerminalTheme))
    {
        return E_INVALIDARG;
    }
    _dpi = std::max(theme->dpi, 1u);
    _themeDark = theme->dark != 0u;
    _themeHighContrast = theme->highContrast != 0u;
    _backgroundArgb = theme->backgroundArgb;
    _foregroundArgb = theme->foregroundArgb;
    _cursorArgb = theme->cursorArgb;
    _selectionBackgroundArgb = theme->selectionBackgroundArgb;
    _selectionForegroundArgb = theme->selectionForegroundArgb;
    _hyperlinkArgb = theme->hyperlinkArgb;
    discardDeviceResources();
    resizeTerminal();
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        applyNativeScrollbarTheme(hwnd);
        InvalidateRect(hwnd, nullptr, FALSE);
    }
    return S_OK;
}

HRESULT STDMETHODCALLTYPE Terminal::UpdateSourceLocation(const TerminalSourceUpdate* update) noexcept
{
    if (update == nullptr || update->sizeBytes < sizeof(TerminalSourceUpdate) ||
        ! validateLocation(update->sourceLocation) ||
        (update->disposition != TerminalSourceDisposition::Active &&
         update->disposition != TerminalSourceDisposition::Detached))
    {
        return E_INVALIDARG;
    }
    if (update->originalSource.folderWindowInstanceId != _originalSource.folderWindowInstanceId ||
        update->originalSource.paneInstanceId != _originalSource.paneInstanceId)
    {
        return E_INVALIDARG;
    }

    const bool windowsLocation = update->sourceLocation.kind == TerminalLocationKind::WindowsLocal ||
        update->sourceLocation.kind == TerminalLocationKind::WindowsUnc;
    const bool pluginBackedLocation = update->sourceLocation.kind == TerminalLocationKind::PluginBacked &&
        ! copySpan(update->sourceLocation.pluginShortId).empty();
    const bool wslLocation = update->sourceLocation.kind == TerminalLocationKind::Wsl &&
        OrdinalString::EqualsNoCase(copySpan(update->sourceLocation.wslDistribution), _launchWslDistribution);
    const std::wstring target = CopyTerminalLocationPath(update->sourceLocation);
    const bool supportedTarget = ((windowsLocation || pluginBackedLocation) && IsSupportedAbsoluteWindowsPath(target)) ||
        (wslLocation && IsSupportedAbsoluteWslPath(target));
    const bool followEnabled = _followPathWhenIdle.load(std::memory_order_acquire);
    std::wstring currentPromptPath;
    std::wstring currentHistoryPath;
    std::wstring currentPromptBuffer;
    std::vector<std::wstring> currentSessionHistory;
    bool supportsPromptReplacement = false;
    bool applyNow = false;
    {
        std::unique_lock stateLock(_stateMutex);
        if (update->sourceGeneration <= _sourceGeneration)
        {
            return S_FALSE;
        }
        _sourceGeneration = update->sourceGeneration;
        if (update->disposition == TerminalSourceDisposition::Detached)
        {
            _sourceLocationKind = TerminalLocationKind::Unsupported;
            _sourceLocationPath.clear();
            _sourcePluginShortId.clear();
            _pendingFollowDirectory.clear();
            _followState = TerminalFollowState::Detached;
            _stateGeneration.fetch_add(1u, std::memory_order_acq_rel);
            stateLock.unlock();
            return S_OK;
        }
        _sourceLocationKind = supportedTarget ? update->sourceLocation.kind : TerminalLocationKind::Unsupported;
        _sourceLocationPath = supportedTarget ? target : std::wstring{};
        _sourcePluginShortId = supportedTarget ? copySpan(update->sourceLocation.pluginShortId) : std::wstring{};
        if (! followEnabled || ! supportedTarget || (! windowsLocation && ! pluginBackedLocation) ||
            _shellKind != ShellKind::PowerShell)
        {
            _pendingFollowDirectory.clear();
            _followState = followEnabled ? TerminalFollowState::Paused : TerminalFollowState::Disabled;
            _stateGeneration.fetch_add(1u, std::memory_order_acq_rel);
            stateLock.unlock();
            return S_FALSE;
        }
        _pendingFollowDirectory = target;
        _followState = TerminalFollowState::Pending;
        applyNow = _integrationTrusted && _idleAtPrimaryPrompt && ! _hasPendingUserInput;
        if (applyNow)
        {
            currentPromptPath = _trustedCurrentDirectory;
            currentHistoryPath = _trustedHistoryPath;
            currentPromptBuffer = _trustedPromptBuffer;
            currentSessionHistory = _trustedCurrentSessionHistory;
            supportsPromptReplacement = _trustedHistoryProviderAvailable;
        }
        _stateGeneration.fetch_add(1u, std::memory_order_acq_rel);
    }
    if (applyNow)
    {
        static_cast<void>(acceptTrustedPrompt(std::move(currentPromptPath),
                                              std::move(currentHistoryPath),
                                              std::move(currentPromptBuffer),
                                              std::move(currentSessionHistory),
                                              supportsPromptReplacement));
        return S_OK;
    }
    return S_FALSE;
}



HRESULT Terminal::copyOwned(std::wstring_view text, TerminalOwnedUtf16& output) noexcept
{
    output = {};
    if (text.empty())
    {
        return S_OK;
    }
    if (text.size() > (std::numeric_limits<uint32_t>::max)() || text.size() > ((std::numeric_limits<size_t>::max)() / sizeof(wchar_t)) - 1u)
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }
    const size_t bytes = (text.size() + 1u) * sizeof(wchar_t);
    auto* data = static_cast<wchar_t*>(CoTaskMemAlloc(bytes));
    if (data == nullptr)
    {
        return E_OUTOFMEMORY;
    }
    memcpy(data, text.data(), text.size() * sizeof(wchar_t));
    data[text.size()] = L'\0';
    output.data = data;
    output.length = static_cast<uint32_t>(text.size());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE Terminal::GetViewState(TerminalViewState* state) noexcept
{
    if (state == nullptr || state->sizeBytes < sizeof(TerminalViewState))
    {
        return E_INVALIDARG;
    }
    TerminalViewState result{};
    result.sizeBytes = sizeof(result);
    result.instanceId = _instanceId;
    result.stateGeneration = _stateGeneration.load(std::memory_order_acquire);
    result.sessionGeneration = _sessionGeneration.load(std::memory_order_acquire);
    result.capabilityFlags = TerminalCapabilityPathInsertion;
    result.exitCodePresent = _exitCodePresent.load(std::memory_order_acquire) ? 1u : 0u;
    result.exitCode = _exitCode.load(std::memory_order_acquire);
    result.finalSnapshotComplete = _finalSnapshotComplete.load(std::memory_order_acquire) ? 1u : 0u;
    result.activity.sizeBytes = sizeof(result.activity);
    result.activity.lifecycleState = _lifecycle.load(std::memory_order_acquire);
    result.activity.generation = result.stateGeneration;
    {
        std::scoped_lock stateLock(_stateMutex);
        result.followState = _followState;
        result.activity.activityTrust = _integrationTrusted ? TerminalActivityTrust::Trusted : TerminalActivityTrust::Untrusted;
        result.activity.idleAtPrimaryPrompt = _integrationTrusted && _idleAtPrimaryPrompt ? 1u : 0u;
        result.activity.hasRunningCommandOrChild = _integrationTrusted && ! _idleAtPrimaryPrompt ? 1u : 0u;
        result.activity.hasPendingUserInput = _integrationTrusted && _hasPendingUserInput ? 1u : 0u;
        if (_integrationTrusted)
        {
            result.capabilityFlags |= TerminalCapabilityPromptTracking;
            if (_followPathWhenIdle.load(std::memory_order_acquire))
            {
                result.capabilityFlags |= TerminalCapabilityFollow;
            }
        }
    }

    const std::wstring title = LoadStringResource(g_hInstance, IDS_TERMINAL_NAME);
    std::wstring status;
    switch (result.activity.lifecycleState)
    {
    case TerminalLifecycleState::Running:
        status = LoadStringResource(g_hInstance, IDS_TERMINAL_STATUS_RUNNING);
#if defined(ENABLE_TESTS)
        if (result.activity.activityTrust == TerminalActivityTrust::Untrusted && _shellKind == ShellKind::PowerShell)
        {
            status.append(FormatStringResource(g_hInstance,
                                               IDS_TERMINAL_STATUS_INTEGRATION_ERROR,
                                               _integrationDebugStage.load(std::memory_order_acquire),
                                               static_cast<uint32_t>(_integrationDebugResult.load(std::memory_order_acquire)),
                                               _integrationDebugClientPid.load(std::memory_order_acquire)));
        }
#endif
        break;
    case TerminalLifecycleState::Diagnostic:
    case TerminalLifecycleState::Failed:
        status = LoadStringResource(g_hInstance, IDS_TERMINAL_STATUS_DIAGNOSTIC);
        break;
    case TerminalLifecycleState::Exited:
        status = LoadStringResource(g_hInstance, IDS_TERMINAL_STATUS_EXITED);
        break;
    case TerminalLifecycleState::Created:
    case TerminalLifecycleState::Starting:
    case TerminalLifecycleState::DrainingAfterRootExit:
    case TerminalLifecycleState::Closing:
    case TerminalLifecycleState::Retiring:
    case TerminalLifecycleState::Quarantined:
        status = LoadStringResource(g_hInstance, IDS_TERMINAL_STATUS_STARTING);
        break;
    }
    HRESULT hr = copyOwned(title, result.title);
    if (SUCCEEDED(hr))
    {
        hr = copyOwned(status, result.status);
    }
    if (FAILED(hr))
    {
        CoTaskMemFree(result.title.data);
        CoTaskMemFree(result.status.data);
        return hr;
    }
    *state = result;
    return S_OK;
}

void Terminal::recordShortcutRoute(TerminalShortcutRoute route,
                                   HRESULT result,
                                   std::chrono::steady_clock::time_point startedAt) noexcept
{
    if (startedAt == std::chrono::steady_clock::time_point{})
    {
        return;
    }
    const uint64_t durationUs = Debug::Perf::ElapsedUs(startedAt);
    ++_shortcutRouteBatchCount;
    _shortcutRouteBatchTotalUs += durationUs;
    _shortcutRouteBatchMaximumUs = std::max(_shortcutRouteBatchMaximumUs, durationUs);
    if (FAILED(result))
    {
        ++_shortcutRouteBatchFailed;
    }
    switch (route)
    {
        case TerminalShortcutRoute::PassThrough: ++_shortcutRouteBatchPassThrough; break;
        case TerminalShortcutRoute::Handled: ++_shortcutRouteBatchHandled; break;
        case TerminalShortcutRoute::InvokeHostCommand: ++_shortcutRouteBatchHost; break;
        case TerminalShortcutRoute::Blocked: ++_shortcutRouteBatchBlocked; break;
    }
    if (_shortcutRouteBatchCount >= 250u)
    {
        flushShortcutRouteMetrics();
    }
}

void Terminal::flushShortcutRouteMetrics() noexcept
{
    if (_shortcutRouteBatchCount == 0u)
    {
        return;
    }
    Debug::Perf::EmitDurationUs(L"terminal.shortcut.route_batch_us",
                                _shortcutRouteBatchTotalUs,
                                _shortcutRouteBatchCount,
                                _shortcutRouteBatchMaximumUs);
    Debug::Perf::Emit(L"terminal.shortcut.route_counts",
                      L"handled-passThrough",
                      0u,
                      _shortcutRouteBatchHandled,
                      _shortcutRouteBatchPassThrough);
    Debug::Perf::Emit(L"terminal.shortcut.route_counts",
                      L"host-blocked",
                      0u,
                      _shortcutRouteBatchHost,
                      _shortcutRouteBatchBlocked);
    Debug::Perf::EmitValue(L"terminal.shortcut.route_failed", _shortcutRouteBatchFailed);
    Debug::Perf::EmitValue(L"terminal.shortcut.route_allocations", 0u);
    _shortcutRouteBatchCount = 0u;
    _shortcutRouteBatchTotalUs = 0u;
    _shortcutRouteBatchMaximumUs = 0u;
    _shortcutRouteBatchHandled = 0u;
    _shortcutRouteBatchPassThrough = 0u;
    _shortcutRouteBatchHost = 0u;
    _shortcutRouteBatchBlocked = 0u;
    _shortcutRouteBatchFailed = 0u;
}

HRESULT STDMETHODCALLTYPE Terminal::RouteShortcut(
    const TerminalShortcutRequest* request, TerminalShortcutRoute* route) noexcept
{
    if (route == nullptr)
    {
        return E_POINTER;
    }
    const auto startedAt = Debug::Perf::IsCaptureEnabled()
        ? std::chrono::steady_clock::now()
        : std::chrono::steady_clock::time_point{};
    HRESULT result = E_UNEXPECTED;
    const auto recordRoute = wil::scope_exit([&]() noexcept { recordShortcutRoute(*route, result, startedAt); });
    *route = TerminalShortcutRoute::PassThrough;
    if (request == nullptr || request->sizeBytes < sizeof(TerminalShortcutRequest) || ! IsSpanWellFormed(request->commandId) ||
        (request->message != WM_KEYDOWN && request->message != WM_SYSKEYDOWN) ||
        memcmp(request->instanceId.bytes, _instanceId.bytes, sizeof(_instanceId.bytes)) != 0 ||
        request->sessionGeneration != _sessionGeneration.load(std::memory_order_acquire))
    {
        result = E_INVALIDARG;
        return result;
    }

    if (confirmationVisible())
    {
        *route = TerminalShortcutRoute::Blocked;
        result = S_OK;
        return result;
    }

    const bool altGr = (request->modifierFlags & TerminalShortcutModifierRightAlt) != 0u &&
        (request->modifierFlags & (TerminalShortcutModifierCtrl | TerminalShortcutModifierAlt)) ==
            (TerminalShortcutModifierCtrl | TerminalShortcutModifierAlt);
    const bool printableKey = (request->virtualKey >= static_cast<uint32_t>('0') && request->virtualKey <= static_cast<uint32_t>('Z')) ||
        (request->virtualKey >= VK_OEM_1 && request->virtualKey <= VK_OEM_8) || request->virtualKey == VK_OEM_102;
    if (altGr && printableKey)
    {
        *route = TerminalShortcutRoute::PassThrough;
        result = S_OK;
        return result;
    }

    const std::wstring_view commandId = ViewTerminalSpan(request->commandId);
    if (commandId == L"cmd/terminal/copySelectionOrPassthrough")
    {
        const bool selectionPresent = hasSelection();
        if (selectionPresent)
        {
            static_cast<void>(copySelection());
            *route = TerminalShortcutRoute::Handled;
        }
        else
        {
            *route = TerminalShortcutRoute::PassThrough;
        }
        result = S_OK;
        return result;
    }

    if (! IsTerminalPluginAction(commandId))
    {
        *route = TerminalShortcutRoute::InvokeHostCommand;
        result = S_OK;
        return result;
    }

    TerminalActionRequest action{};
    action.sizeBytes = sizeof(action);
    action.commandId = request->commandId;
    action.instanceId = request->instanceId;
    action.sessionGeneration = request->sessionGeneration;
    const HRESULT actionHr = ExecuteAction(&action);
    if (actionHr == E_NOTIMPL)
    {
        *route = TerminalShortcutRoute::PassThrough;
        result = S_OK;
        return result;
    }
    if (FAILED(actionHr))
    {
        result = actionHr;
        return result;
    }
    *route = TerminalShortcutRoute::Handled;
    result = S_OK;
    return result;
}

HRESULT STDMETHODCALLTYPE Terminal::GetActionState(
    const TerminalActionRequest* request, TerminalActionState* state) noexcept
{
    if (state == nullptr)
    {
        return E_POINTER;
    }
    const uint32_t callerSizeBytes = state->sizeBytes;
    if (callerSizeBytes < sizeof(TerminalActionState))
    {
        return E_INVALIDARG;
    }
    if (request == nullptr || request->sizeBytes < sizeof(TerminalActionRequest) || ! IsSpanWellFormed(request->commandId) ||
        memcmp(request->instanceId.bytes, _instanceId.bytes, sizeof(_instanceId.bytes)) != 0 ||
        request->sessionGeneration != _sessionGeneration.load(std::memory_order_acquire))
    {
        return E_INVALIDARG;
    }

    TerminalActionState result{};
    result.sizeBytes = sizeof(result);
    const std::wstring_view commandId = ViewTerminalSpan(request->commandId);
    if (! IsTerminalPluginAction(commandId))
    {
        return E_NOTIMPL;
    }
    const bool surfaceAction = IsTerminalCommandSurfaceAction(commandId);
    const bool runningSurface = _lifecycle.load(std::memory_order_acquire) == TerminalLifecycleState::Running ||
        (commandId == L"cmd/terminal/find" && _finalSnapshotComplete.load(std::memory_order_acquire));
    const bool copyAction = commandId == L"cmd/terminal/copy" || commandId == L"cmd/terminal/copySelectionOrPassthrough";
    result.enabled = ! confirmationVisible() && (! copyAction || hasSelection()) && (! surfaceAction || runningSurface)
        ? 1u
        : 0u;
    if (const HWND hwnd = _windowHandle.load(std::memory_order_acquire); hwnd != nullptr)
    {
        GetClientRect(hwnd, &result.childAnchor);
        result.anchorPresent = 1u;
    }
    *state = result;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE Terminal::ExecuteAction(const TerminalActionRequest* request) noexcept
{
    if (request == nullptr || request->sizeBytes < sizeof(TerminalActionRequest) || ! IsSpanWellFormed(request->commandId) ||
        memcmp(request->instanceId.bytes, _instanceId.bytes, sizeof(_instanceId.bytes)) != 0 ||
        request->sessionGeneration != _sessionGeneration.load(std::memory_order_acquire))
    {
        return E_INVALIDARG;
    }
    if (confirmationVisible())
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DISABLED_BY_POLICY);
    }

    const std::wstring_view commandId = ViewTerminalSpan(request->commandId);
    if (! IsTerminalPluginAction(commandId))
    {
        return E_NOTIMPL;
    }
    if (IsTerminalCommandSurfaceAction(commandId))
    {
        const bool available = _lifecycle.load(std::memory_order_acquire) == TerminalLifecycleState::Running ||
            (commandId == L"cmd/terminal/find" && _finalSnapshotComplete.load(std::memory_order_acquire));
        if (! available)
        {
            return E_NOTIMPL;
        }
    }
    if (commandId == L"cmd/terminal/copy")
    {
        return copySelection() ? S_OK : S_FALSE;
    }
    if (commandId == L"cmd/terminal/copySelectionOrPassthrough")
    {
        // Palette/host execute must never pass Enter/ETX. Copy when a selection
        // exists; otherwise return a no-op without writing input.
        if (! hasSelection())
        {
            return S_FALSE;
        }
        return copySelection() ? S_OK : S_FALSE;
    }
    if (commandId == L"cmd/terminal/paste")
    {
        return pasteClipboard() ? S_OK : S_FALSE;
    }
    if (commandId == L"cmd/terminal/selectAll")
    {
        return selectAll() ? S_OK : S_FALSE;
    }
    if (commandId == L"cmd/terminal/scroll/lineDown")
    {
        scrollViewport(1);
        return S_OK;
    }
    if (commandId == L"cmd/terminal/scroll/lineUp")
    {
        scrollViewport(-1);
        return S_OK;
    }
    if (commandId == L"cmd/terminal/scroll/pageDown")
    {
        scrollViewport(static_cast<intptr_t>(std::max<uint16_t>(_rows, 1u)));
        return S_OK;
    }
    if (commandId == L"cmd/terminal/scroll/pageUp")
    {
        scrollViewport(-static_cast<intptr_t>(std::max<uint16_t>(_rows, 1u)));
        return S_OK;
    }
    if (commandId == L"cmd/terminal/scroll/top")
    {
        scrollViewportTo(0u);
        return S_OK;
    }
    if (commandId == L"cmd/terminal/scroll/bottom")
    {
        scrollViewportTo((std::numeric_limits<uint64_t>::max)());
        return S_OK;
    }
    if (commandId == L"cmd/terminal/font/increase")
    {
        setLiveFontSize(_liveFontSizeDip + 1.0f);
        return S_OK;
    }
    if (commandId == L"cmd/terminal/font/decrease")
    {
        setLiveFontSize(_liveFontSizeDip - 1.0f);
        return S_OK;
    }
    if (commandId == L"cmd/terminal/font/reset")
    {
        setLiveFontSize(_config.fontSizeDip);
        return S_OK;
    }
    if (commandId == L"cmd/terminal/find")
    {
        return openCommandSurface(CommandSurfaceKind::Find) ? S_OK : E_NOTIMPL;
    }
    if (commandId == L"cmd/terminal/suggestions")
    {
        return openCommandSurface(CommandSurfaceKind::Suggestions) ? S_OK : E_NOTIMPL;
    }
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE Terminal::SetCallback(ITerminalEventCallback* callback, void* cookie) noexcept
{
    if ((callback == nullptr) != (cookie == nullptr))
    {
        return E_INVALIDARG;
    }
    {
        // publishFinalExitEvent holds the same mutex across callback dispatch,
        // making callback removal a hard no-in-flight/no-future barrier.
        std::scoped_lock callbackLock(_eventCallbackMutex);
        _eventCallback = callback;
        _eventCallbackCookie = cookie;
    }
    if (callback != nullptr && _finalSnapshotComplete.load(std::memory_order_acquire))
    {
        notifyOutputReady();
    }
    return S_OK;
}

HRESULT STDMETHODCALLTYPE Terminal::Close() noexcept
{
    if (_closing.exchange(true, std::memory_order_acq_rel))
    {
        return S_FALSE;
    }
    prepareClose();

    if (_closeWork)
    {
        AddRef();
        SubmitThreadpoolWork(_closeWork.get());
        return S_OK;
    }

    finishClose();
    return S_OK;
}






Terminal::MessageRouteResult Terminal::routeInputMessage(
    HWND, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    switch (message)
    {
    case WM_GETDLGCODE:
        return {true, DLGC_WANTALLKEYS | DLGC_WANTCHARS | DLGC_WANTARROWS};
    case WM_KEYDOWN:
    case WM_SYSKEYDOWN:
        if (confirmationVisible())
        {
            if (wParam == VK_RETURN)
            {
                MarkTranslatedCharacterConsumed(_translatedCharacterSuppressionPending);
                resolveConfirmation(true);
            }
            else if (wParam == VK_ESCAPE)
            {
                MarkTranslatedCharacterConsumed(_translatedCharacterSuppressionPending);
                resolveConfirmation(false);
            }
            return {true, 0};
        }
        if (_commandSurfaceKind != CommandSurfaceKind::None)
        {
            static_cast<void>(handleCommandSurfaceKey(wParam));
            return {true, 0};
        }
        if (encodeSpecialKey(wParam, lParam, false))
        {
            return {true, 0};
        }
        return {};
    case WM_KEYUP:
    case WM_SYSKEYUP:
        if (confirmationVisible())
        {
            return {true, 0};
        }
        if (_commandSurfaceKind != CommandSurfaceKind::None)
        {
            return {true, 0};
        }
        if (encodeSpecialKey(wParam, lParam, true))
        {
            return {true, 0};
        }
        return {};
    case WM_CHAR:
    {
        if (ConsumeTranslatedCharacterSuppression(_translatedCharacterSuppressionPending))
        {
            return {true, 0};
        }
        if (confirmationVisible())
        {
            return {true, 0};
        }
        wchar_t character = static_cast<wchar_t>(wParam);
        if (handleCommandSurfaceCharacter(character))
        {
            return {true, 0};
        }
        if (character == L'\b')
        {
            if (writeInput("\x7f"))
            {
                markUserInput("\x7f");
            }
            return {true, 0};
        }
        if (character == L'\n')
        {
            return {true, 0};
        }
        if (IS_HIGH_SURROGATE(character))
        {
            _pendingHighSurrogate = character;
            return {true, 0};
        }
        std::array<wchar_t, 2> text{};
        size_t length = 1u;
        if (IS_LOW_SURROGATE(character) && _pendingHighSurrogate != L'\0')
        {
            text[0] = _pendingHighSurrogate;
            text[1] = character;
            length = 2u;
        }
        else
        {
            text[0] = character;
        }
        _pendingHighSurrogate = L'\0';
        static_cast<void>(writeTextInput(std::wstring_view(text.data(), length)));
        return {true, 0};
    }
    case WM_SYSCHAR:
    {
        const bool suppressCharacter = ConsumeTranslatedCharacterSuppression(_translatedCharacterSuppressionPending);
        return confirmationVisible() || _commandSurfaceKind != CommandSurfaceKind::None || suppressCharacter
            ? MessageRouteResult{true, 0}
            : MessageRouteResult{};
    }
    default:
        return {};
    }
}

bool Terminal::handleTimer(HWND hwnd, UINT_PTR timerId) noexcept
{
    if (timerId == kKittyUploadRetryTimer)
    {
        KillTimer(hwnd, kKittyUploadRetryTimer);
        _kittyNextUploadRetryAtTick = 0u;
        requestKittyRepaint(hwnd);
        return true;
    }
    if (timerId == kSelectionAutoScrollTimer && _selecting && _selectionAutoScrollRows != 0)
    {
        scrollViewport(_selectionAutoScrollRows);
        static_cast<void>(updateSelection(_selectionLastPoint));
        return true;
    }
    return false;
}

Terminal::MessageRouteResult Terminal::routeMouseMessage(
    HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    switch (message)
    {
    case WM_LBUTTONDOWN:
    {
        SetFocus(hwnd);
        const POINT point{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
        if (handleConfirmationClick(point))
        {
            return {true, 0};
        }
        if (handleCommandSurfaceClick(point, false))
        {
            return {true, 0};
        }
        if ((GetKeyState(VK_CONTROL) & 0x8000) != 0 && activateHyperlinkAt(point))
        {
            return {true, 0};
        }
        if ((GetKeyState(VK_SHIFT) & 0x8000) == 0 &&
            ! HasReportedMouseButton(_reportedMouseButtons, GHOSTTY_MOUSE_BUTTON_LEFT) &&
            encodeMouse(GHOSTTY_MOUSE_ACTION_PRESS, GHOSTTY_MOUSE_BUTTON_LEFT, point, true))
        {
            static_cast<void>(AddReportedMouseButton(_reportedMouseButtons, GHOSTTY_MOUSE_BUTTON_LEFT));
            _reportedMouseLastPoint = point;
            SetCapture(hwnd);
            return {true, 0};
        }
        stopSelection();
        bool selectionStarted = false;
        {
            std::scoped_lock lock(_terminalMutex);
            if (_ghosttyTerminal != nullptr && _runtime.terminalGridRefTrack != nullptr &&
                _runtime.terminalGridRefTrack(
                    _ghosttyTerminal, terminalPointFromClient(point), &_selectionAnchor) == GHOSTTY_SUCCESS &&
                _selectionAnchor != nullptr)
            {
                static_cast<void>(_runtime.terminalSet(_ghosttyTerminal, GHOSTTY_TERMINAL_OPT_SELECTION, nullptr));
                selectionStarted = true;
            }
        }
        _selectionLastPoint = point;
        _selecting = selectionStarted;
        if (selectionStarted)
        {
            SetCapture(hwnd);
        }
        InvalidateRect(hwnd, nullptr, FALSE);
        return {true, 0};
    }
    case WM_LBUTTONDBLCLK:
        SetFocus(hwnd);
        if (handleConfirmationClick(POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)}))
        {
            return {true, 0};
        }
        if (handleCommandSurfaceClick(POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)}, true))
        {
            return {true, 0};
        }
        if ((GetKeyState(VK_SHIFT) & 0x8000) == 0 &&
            ! HasReportedMouseButton(_reportedMouseButtons, GHOSTTY_MOUSE_BUTTON_LEFT) &&
            encodeMouse(
                GHOSTTY_MOUSE_ACTION_PRESS,
                GHOSTTY_MOUSE_BUTTON_LEFT,
                POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)},
                true))
        {
            static_cast<void>(AddReportedMouseButton(_reportedMouseButtons, GHOSTTY_MOUSE_BUTTON_LEFT));
            _reportedMouseLastPoint = POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            SetCapture(hwnd);
            return {true, 0};
        }
        stopSelection();
        if (GetCapture() == hwnd)
        {
            ReleaseCapture();
        }
        static_cast<void>(selectWordAt(POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)}));
        return {true, 0};
    case WM_MOUSEMOVE:
        if (_commandSurfaceKind != CommandSurfaceKind::None)
        {
            return {true, 0};
        }
        if (confirmationVisible())
        {
            return {true, 0};
        }
        if (_reportedMouseButtons != 0u)
        {
            _reportedMouseLastPoint = POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            static_cast<void>(encodeMouse(
                GHOSTTY_MOUSE_ACTION_MOTION,
                GHOSTTY_MOUSE_BUTTON_UNKNOWN,
                POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)},
                true));
            return {true, 0};
        }
        if (_selecting && (wParam & MK_LBUTTON) != 0u)
        {
            const POINT point{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            _selectionLastPoint = point;
            RECT client{};
            GetClientRect(hwnd, &client);
            _selectionAutoScrollRows = point.y < client.top ? -1 : point.y >= client.bottom ? 1 : 0;
            if (_selectionAutoScrollRows != 0)
            {
                static_cast<void>(SetTimer(hwnd, kSelectionAutoScrollTimer, 50u, nullptr));
            }
            else
            {
                KillTimer(hwnd, kSelectionAutoScrollTimer);
            }
            static_cast<void>(updateSelection(point));
            return {true, 0};
        }
        if ((GetKeyState(VK_SHIFT) & 0x8000) == 0 &&
            encodeMouse(
                GHOSTTY_MOUSE_ACTION_MOTION,
                GHOSTTY_MOUSE_BUTTON_UNKNOWN,
                POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)},
                false))
        {
            return {true, 0};
        }
        return {};
    case WM_LBUTTONUP:
        if (_commandSurfaceKind != CommandSurfaceKind::None)
        {
            return {true, 0};
        }
        if (confirmationVisible())
        {
            return {true, 0};
        }
        if (RemoveReportedMouseButton(_reportedMouseButtons, GHOSTTY_MOUSE_BUTTON_LEFT))
        {
            static_cast<void>(encodeMouse(
                GHOSTTY_MOUSE_ACTION_RELEASE,
                GHOSTTY_MOUSE_BUTTON_LEFT,
                POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)},
                _reportedMouseButtons != 0u));
            if (_reportedMouseButtons == 0u && GetCapture() == hwnd)
            {
                ReleaseCapture();
            }
            return {true, 0};
        }
        if (_selecting)
        {
            static_cast<void>(updateSelection(POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)}));
            stopSelection();
            if (GetCapture() == hwnd)
            {
                ReleaseCapture();
            }
        }
        return {true, 0};
    case WM_CAPTURECHANGED:
        stopSelection();
        releaseReportedMouseButtons();
        return {true, 0};
    case WM_TIMER:
        return {handleTimer(hwnd, static_cast<UINT_PTR>(wParam)), 0};
    case WM_MBUTTONDOWN:
    case WM_RBUTTONDOWN:
    {
        if (confirmationVisible())
        {
            return {true, 0};
        }
        if (_commandSurfaceKind != CommandSurfaceKind::None)
        {
            closeCommandSurface();
            return {true, 0};
        }
        SetFocus(hwnd);
        const GhosttyMouseButton button =
            message == WM_MBUTTONDOWN ? GHOSTTY_MOUSE_BUTTON_MIDDLE : GHOSTTY_MOUSE_BUTTON_RIGHT;
        if ((GetKeyState(VK_SHIFT) & 0x8000) == 0 && ! HasReportedMouseButton(_reportedMouseButtons, button) &&
            encodeMouse(
                GHOSTTY_MOUSE_ACTION_PRESS, button, POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)}, true))
        {
            static_cast<void>(AddReportedMouseButton(_reportedMouseButtons, button));
            _reportedMouseLastPoint = POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            SetCapture(hwnd);
            return {true, 0};
        }
        return {};
    }
    case WM_MBUTTONUP:
    case WM_RBUTTONUP:
        if (_commandSurfaceKind != CommandSurfaceKind::None)
        {
            return {true, 0};
        }
        if (confirmationVisible())
        {
            return {true, 0};
        }
        {
            const GhosttyMouseButton button =
                message == WM_MBUTTONUP ? GHOSTTY_MOUSE_BUTTON_MIDDLE : GHOSTTY_MOUSE_BUTTON_RIGHT;
            if (! RemoveReportedMouseButton(_reportedMouseButtons, button))
            {
                return {};
            }

            static_cast<void>(encodeMouse(
                GHOSTTY_MOUSE_ACTION_RELEASE,
                button,
                POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)},
                _reportedMouseButtons != 0u));
            if (_reportedMouseButtons == 0u && GetCapture() == hwnd)
            {
                ReleaseCapture();
            }
            return {true, 0};
        }
    case WM_MOUSEWHEEL:
    {
        if (confirmationVisible())
        {
            return {true, 0};
        }
        if (_commandSurfaceKind != CommandSurfaceKind::None)
        {
            navigateCommandSurface(GET_WHEEL_DELTA_WPARAM(wParam) > 0 ? -1 : 1);
            return {true, 0};
        }
        POINT point{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
        ScreenToClient(hwnd, &point);
        const short delta = GET_WHEEL_DELTA_WPARAM(wParam);
        const GhosttyMouseButton button = delta > 0 ? GHOSTTY_MOUSE_BUTTON_FOUR : GHOSTTY_MOUSE_BUTTON_FIVE;
        if ((GetKeyState(VK_SHIFT) & 0x8000) == 0 && encodeMouse(GHOSTTY_MOUSE_ACTION_PRESS, button, point, false))
        {
            return {true, 0};
        }
        UINT linesPerDetent = 3u;
        static_cast<void>(SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0u, &linesPerDetent, 0u));
        const TerminalWheelScrollResult scroll = ResolveTerminalWheelScroll(
            _mouseWheelDeltaRemainder, static_cast<int>(delta), linesPerDetent, _rows);
        _mouseWheelDeltaRemainder = scroll.remainder;
        if (scroll.rows != 0)
        {
            scrollViewport(scroll.rows);
        }
        return {true, 0};
    }
    case WM_VSCROLL:
        handleVerticalScroll(LOWORD(wParam));
        return {true, 0};
    default:
        return {};
    }
}

Terminal::MessageRouteResult Terminal::routeImeMessage(HWND hwnd, UINT message, WPARAM, LPARAM lParam) noexcept
{
    switch (message)
    {
    case WM_IME_STARTCOMPOSITION:
        _imeComposition.clear();
        updateImeCandidatePosition();
        return {true, 0};
    case WM_IME_COMPOSITION:
        if ((lParam & GCS_RESULTSTR) != 0)
        {
            const std::optional<std::wstring> result = ReadImeString(hwnd, GCS_RESULTSTR);
            if (result.has_value())
            {
                if (_commandSurfaceKind != CommandSurfaceKind::None)
                {
                    for (const wchar_t character : result.value())
                    {
                        static_cast<void>(handleCommandSurfaceCharacter(character));
                    }
                }
                else
                {
                    static_cast<void>(writeTextInput(result.value()));
                }
            }
            _imeComposition.clear();
        }
        else if ((lParam & GCS_COMPSTR) != 0)
        {
            _imeComposition = ReadImeString(hwnd, GCS_COMPSTR).value_or(std::wstring{});
        }
        updateImeCandidatePosition();
        InvalidateRect(hwnd, nullptr, FALSE);
        return {true, 0};
    case WM_IME_ENDCOMPOSITION:
        _imeComposition.clear();
        InvalidateRect(hwnd, nullptr, FALSE);
        return {true, 0};
    default:
        return {};
    }
}

Terminal::MessageRouteResult Terminal::routeAccessibilityMessage(
    HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    switch (message)
    {
    case WM_GETOBJECT:
        consumeOutputReady();
        publishAccessibilitySnapshot();
        return {true, _accessibility ? _accessibility->HandleGetObject(wParam, lParam) : 0};
    case WM_SETFOCUS:
        publishAccessibilitySnapshot();
        InvalidateRect(hwnd, nullptr, FALSE);
        return {true, 0};
    case WM_KILLFOCUS:
        if (_commandSurfaceKind == CommandSurfaceKind::Suggestions && ! _commandSurfaceInsertionPending)
        {
            closeCommandSurface();
        }
        publishAccessibilitySnapshot();
        InvalidateRect(hwnd, nullptr, FALSE);
        return {true, 0};
    default:
        return {};
    }
}

Terminal::MessageRouteResult Terminal::routeSessionMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    switch (message)
    {
    case WndMsg::kTerminalOutputReady:
    {
        std::unique_ptr<ReadyWakePayload> payload = TakeMessagePayload<ReadyWakePayload>(lParam);
        if (! payload || _closing.load(std::memory_order_acquire) ||
            payload->sessionGeneration != _sessionGeneration.load(std::memory_order_acquire))
        {
            return {true, 0};
        }
        consumeOutputReady();
        updateScrollbar();
        if (UiaClientsAreListening() != FALSE)
        {
            publishAccessibilitySnapshot();
        }
        InvalidateRect(hwnd, nullptr, FALSE);
        return {true, 0};
    }
    case WndMsg::kTerminalKittyReady:
    {
        std::unique_ptr<ReadyWakePayload> payload = TakeMessagePayload<ReadyWakePayload>(lParam);
        if (! payload || _closing.load(std::memory_order_acquire) ||
            payload->sessionGeneration != _sessionGeneration.load(std::memory_order_acquire))
        {
            return {true, 0};
        }
        InvalidateRect(hwnd, nullptr, FALSE);
        return {true, 0};
    }
    case WndMsg::kTerminalSuggestionInsertionComplete:
        if (_commandSurfaceKind == CommandSurfaceKind::Suggestions && _commandSurfaceInsertionPending)
        {
            _commandSurfaceInsertionPending = false;
            _commandSurfaceLoading = false;
            if (wParam != FALSE)
            {
                closeCommandSurface();
            }
            else
            {
                _commandSurfaceStatusOverride =
                    LoadStringResource(g_hInstance, IDS_TERMINAL_HISTORY_PROMPT_CHANGED);
                publishAccessibilitySnapshot();
                InvalidateRect(hwnd, nullptr, FALSE);
            }
        }
        return {true, 0};
    case WndMsg::kTerminalCommandSurfaceReady:
    {
        std::unique_ptr<TerminalCommandSurfaceReadyPayload> payload =
            TakeMessagePayload<TerminalCommandSurfaceReadyPayload>(lParam);
        if (payload && payload->snapshotOnly)
        {
            if (_commandSurfaceKind != CommandSurfaceKind::Find ||
                payload->snapshotGeneration != _commandSurfaceSnapshotGeneration ||
                payload->sessionGeneration != _sessionGeneration.load(std::memory_order_acquire))
            {
                if (payload->snapshot)
                {
                    SecureZeroMemory(payload->snapshot->data(), payload->snapshot->size() * sizeof(wchar_t));
                }
                return {true, 0};
            }
            _commandSurfaceSnapshot = std::move(payload->snapshot);
            _commandSurfaceSnapshotReady = true;
            _commandSurfaceLoading = false;
            Debug::Perf::EmitDurationUs(L"terminal.find.snapshot_batch_us",
                                        payload->snapshotDurationUs,
                                        _commandSurfaceSnapshot ? _commandSurfaceSnapshot->size() * sizeof(wchar_t) : 0u,
                                        1u);
            restartCommandSurfaceQuery();
            publishAccessibilitySnapshot();
            InvalidateRect(hwnd, nullptr, FALSE);
            return {true, 0};
        }
        if (! payload || payload->generation != _commandSurfaceGeneration ||
            payload->sessionGeneration != _sessionGeneration.load(std::memory_order_acquire) ||
            payload->suggestions != (_commandSurfaceKind == CommandSurfaceKind::Suggestions))
        {
            return {true, 0};
        }
        _commandSurfaceLoading = false;
        _commandSurfaceSelected = 0u;
        if (payload->suggestions)
        {
            _persistedHistory = std::move(payload->persistedHistory);
            _commandSurfaceSuggestionRows = std::move(payload->suggestion.rows);
            if (payload->historyLoadDurationUs != 0u)
            {
                Debug::Perf::EmitDurationUs(L"terminal.suggestions.load_batch_us",
                                            payload->historyLoadDurationUs,
                                            _persistedHistory ? _persistedHistory->size() : 0u,
                                            payload->suggestion.sourceBytes);
            }
            Debug::Perf::EmitDurationUs(_commandSurfaceQuery.empty()
                                            ? L"terminal.suggestions.open_to_visible_us"
                                            : L"terminal.suggestions.filter_to_visible_us",
                                        payload->durationUs,
                                        payload->suggestion.sourceEntries,
                                        payload->suggestion.sourceBytes);
        }
        else
        {
            _commandSurfaceFindRows = std::move(payload->find.rows);
            _commandSurfaceFindTotal = payload->find.totalMatches;
            Debug::Perf::EmitDurationUs(L"terminal.find.query_to_visible_us",
                                        payload->durationUs,
                                        payload->find.scannedLines,
                                        payload->find.totalMatches);
        }
        publishAccessibilitySnapshot();
        InvalidateRect(hwnd, nullptr, FALSE);
        return {true, 0};
    }
    default:
        return {};
    }
}

Terminal::MessageRouteResult Terminal::routeClipboardMessage(HWND, UINT message, WPARAM, LPARAM lParam) noexcept
{
    switch (message)
    {
    case WndMsg::kTerminalClipboardWrite:
        handleOsc52Clipboard(TakeMessagePayload<Osc52ClipboardPayload>(lParam));
        return {true, 0};
    case WM_PASTE:
        static_cast<void>(pasteClipboard());
        return {true, 0};
    case WM_COPY:
        static_cast<void>(copySelection());
        return {true, 0};
    default:
        return {};
    }
}

Terminal::MessageRouteResult Terminal::routeRenderMessage(HWND hwnd, UINT message, WPARAM, LPARAM) noexcept
{
    switch (message)
    {
    case WM_ERASEBKGND:
        return {true, 1};
    case WM_PAINT:
        consumeOutputReady();
        updateScrollbar();
        if (UiaClientsAreListening() != FALSE)
        {
            publishAccessibilitySnapshot();
        }
        render();
        return {true, 0};
    case WM_SIZE:
        handleResize(hwnd);
        return {true, 0};
    case WM_DPICHANGED_AFTERPARENT:
        _dpi = GetDpiForWindow(hwnd);
        discardDeviceResources();
        resizeTerminal();
        updateScrollbar();
        InvalidateRect(hwnd, nullptr, FALSE);
        return {true, 0};
    default:
        return {};
    }
}

Terminal::MessageRouteResult Terminal::routeTeardownMessage(
    HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    if (message != WM_NCDESTROY)
    {
        return {};
    }

    // Stop payload adoption first, then retire providers before publishing the destroyed window state.
    static_cast<void>(DrainPostedPayloadsForWindow(hwnd));
    _osc52PostPending.store(false, std::memory_order_release);
    _outputNotificationPending.store(false, std::memory_order_release);
    if (_accessibility)
    {
        _accessibility->Retire();
        _accessibility.reset();
    }
    _windowHandle.store(nullptr, std::memory_order_release);
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
    g_terminalWindowCount.fetch_sub(1u, std::memory_order_acq_rel);

    // WM_NCDESTROY still receives the system default processing after project-owned retirement is complete.
    return {true, DefWindowProcW(hwnd, message, wParam, lParam)};
}

LRESULT CALLBACK Terminal::WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    Terminal* self = reinterpret_cast<Terminal*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (message == WM_NCCREATE)
    {
        const auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
        self = static_cast<Terminal*>(create->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        InitPostedPayloadWindow(hwnd);
        g_terminalWindowCount.fetch_add(1u, std::memory_order_acq_rel);
    }
    if (self == nullptr)
    {
        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

    MessageRouteResult routed = self->routeInputMessage(hwnd, message, wParam, lParam);
    if (routed.handled)
    {
        return routed.result;
    }
    routed = self->routeMouseMessage(hwnd, message, wParam, lParam);
    if (routed.handled)
    {
        return routed.result;
    }
    routed = self->routeImeMessage(hwnd, message, wParam, lParam);
    if (routed.handled)
    {
        return routed.result;
    }
    routed = self->routeAccessibilityMessage(hwnd, message, wParam, lParam);
    if (routed.handled)
    {
        return routed.result;
    }
    routed = self->routeSessionMessage(hwnd, message, wParam, lParam);
    if (routed.handled)
    {
        return routed.result;
    }
    routed = self->routeClipboardMessage(hwnd, message, wParam, lParam);
    if (routed.handled)
    {
        return routed.result;
    }
    routed = self->routeRenderMessage(hwnd, message, wParam, lParam);
    if (routed.handled)
    {
        return routed.result;
    }
    routed = self->routeTeardownMessage(hwnd, message, wParam, lParam);
    if (routed.handled)
    {
        return routed.result;
    }
    return DefWindowProcW(hwnd, message, wParam, lParam);
}

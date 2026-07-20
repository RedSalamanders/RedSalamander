#include "DxUi/DxUi.Typography.h"
#include "DxUiTestHelpers.h"

#include <array>
#include <atomic>
#include <cstdint>
#include <fstream>
#include <future>
#include <string_view>
#include <thread>

namespace
{

void TestDxUiTypographyMapsFontRolesToSegoeUiVariableFamilies()
{
    using namespace RedSalamander::DxUi;
    using namespace RedSalamander::DxUi::Typography;

    const TypographySpec bodySpec       = GetDxUiTypographySpec(FontRole::Body);
    const TypographySpec bodyStrongSpec = GetDxUiTypographySpec(FontRole::BodyStrong);
    const TypographySpec listItemSpec   = GetDxUiTypographySpec(FontRole::ListItem);
    const TypographySpec smallSpec      = GetDxUiTypographySpec(FontRole::Small);
    const TypographySpec headerSpec     = GetDxUiTypographySpec(FontRole::Header);
    const TypographySpec titleLargeSpec = GetDxUiTypographySpec(FontRole::TitleLarge);
    const TypographySpec displaySpec    = GetDxUiTypographySpec(FontRole::Display);
    const TypographySpec iconSpec       = GetDxUiTypographySpec(FontRole::Icon);
    const TypographySpec monoSpec       = GetDxUiTypographySpec(FontRole::Monospace);

    Require(bodySpec.familyName == kSegoeUiVariableTextFamily, "body role uses Segoe UI Variable Text");
    Require(bodyStrongSpec.familyName == kSegoeUiVariableTextFamily, "body-strong role uses Segoe UI Variable Text");
    Require(listItemSpec.familyName == kSegoeUiVariableSmallFamily && listItemSpec.sizeDip == 12.0f, "list-item role uses 12 DIP Segoe UI Variable Small");
    Require(smallSpec.familyName == kSegoeUiVariableSmallFamily, "small role uses Segoe UI Variable Small");
    Require(headerSpec.familyName == kSegoeUiVariableSmallFamily, "header role uses Segoe UI Variable Small");
    Require(titleLargeSpec.familyName == kSegoeUiVariableDisplayFamily, "title-large role uses Segoe UI Variable Display");
    Require(displaySpec.familyName == kSegoeUiVariableDisplayFamily, "display role uses Segoe UI Variable Display");
    Require(iconSpec.familyName == kSegoeFluentIconsFamily, "icon role uses Segoe Fluent Icons");
    Require(monoSpec.familyName == kUiMonospaceFamily, "monospace role uses the shared monospace family");
}

void TestDxUiTypographyMeasurementCachesFormatsAndFamilyResolution()
{
    using namespace RedSalamander::DxUi;

    const std::filesystem::path headerPath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Typography.h";
    std::ifstream input(headerPath);
    Require(input.good(), "Typography header is readable for measurement-cache guard");
    const std::string header((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    Require(header.find("struct TypographyFontFamilyCacheEntry") != std::string::npos, "Typography declares a font-family availability cache entry");
    Require(header.find("struct TypographyTextFormatCacheEntry") != std::string::npos, "Typography declares a measurement text-format cache entry");
    Require(header.find("ResolveCachedFontFamilyName(") != std::string::npos, "Typography routes family fallback through a cached resolver");
    Require(header.find("GetCachedMeasurementTextFormat(") != std::string::npos, "Typography exposes a cached measurement text-format accessor");
    Require(header.find("dxui.typography.family_cache_miss_count") != std::string::npos, "Typography family resolution misses emit a gated perf counter");
    Require(header.find("dxui.typography.text_format_cache_miss_count") != std::string::npos, "Typography text-format cache misses emit a gated perf counter");
    Require(header.find("#include \"Helpers.h\"") == std::string::npos, "Typography public header does not include Common perf/logging internals");
    Require(header.find("Debug::Perf::") == std::string::npos, "Typography public header does not inline-link Common perf internals into plugins");

    const auto requireBlock = [](const std::string& text, std::string_view beginMarker, std::string_view endMarker, const char* description)
    {
        const size_t begin = text.find(beginMarker);
        const size_t end   = text.find(endMarker, begin == std::string::npos ? 0u : begin + beginMarker.size());
        Require(begin != std::string::npos && end != std::string::npos && begin < end, description);
        return text.substr(begin, end - begin);
    };

    const std::string createFormatBlock =
        requireBlock(header, "inline HRESULT CreateTextFormat(", "[[nodiscard]] inline HRESULT CreateTextFormatWithStyle", "CreateTextFormat block is found");
    Require(createFormatBlock.find("ResolveCachedFontFamilyName(") != std::string::npos, "CreateTextFormat uses cached family fallback resolution");
    Require(createFormatBlock.find("IsFontFamilyAvailable(dwriteFactory, preferredFamilyBuffer)") == std::string::npos,
            "CreateTextFormat no longer probes the font collection directly on every call");

    const std::string singleLineBlock = requireBlock(
        header, "inline int MeasureSingleLineTextWidthPx(", "[[nodiscard]] inline int MeasureWrappedTextHeightPx", "single-line measurement block is found");
    Require(singleLineBlock.find("GetCachedMeasurementTextFormat(dwriteFactory, role, false") != std::string::npos,
            "single-line measurement reuses the no-wrap cached text format");
    Require(singleLineBlock.find("CreateTextFormat(dwriteFactory, GetDxUiTypographySpec(role)") == std::string::npos,
            "single-line measurement no longer creates a text format per call");

    const std::string wrappedBlock =
        requireBlock(header, "inline int MeasureWrappedTextHeightPx(", "} // namespace RedSalamander::DxUi::Typography", "wrapped measurement block is found");
    Require(wrappedBlock.find("GetCachedMeasurementTextFormat(dwriteFactory, role, true") != std::string::npos,
            "wrapped measurement reuses the wrap cached text format");
    Require(wrappedBlock.find("CreateTextFormat(dwriteFactory, GetDxUiTypographySpec(role)") == std::string::npos,
            "wrapped measurement no longer creates a text format per call");

    Typography::SetTypographyPerfEmitter(
        [](std::wstring_view metric, std::wstring_view detail, uint64_t durationUs, uint64_t value, uint64_t count, HRESULT hr) noexcept
    { Debug::Perf::Emit(metric, detail, durationUs, value, count, hr); });
    const auto resetTypographyPerfEmitter = wil::scope_exit([]() noexcept { Typography::SetTypographyPerfEmitter(nullptr); });

    const int firstWidth = Typography::MeasureSingleLineTextWidthPx(nullptr, FontRole::BodyStrong, L"Typography cache probe");
    const int nextWidth  = Typography::MeasureSingleLineTextWidthPx(nullptr, FontRole::BodyStrong, L"Typography cache probe");
    Require(firstWidth > 0, "typography single-line measurement still returns a positive width");
    Require(nextWidth == firstWidth, "typography cached single-line measurement remains stable");

    const int wrappedHeight = Typography::MeasureWrappedTextHeightPx(nullptr, FontRole::Body, 96, L"Typography wrapped cache probe wraps here");
    Require(wrappedHeight > 0, "typography wrapped measurement still returns a positive height");
}

void TestConnectionCredentialPromptDestroysWindowOnModalQuit()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"RedSalamander" / L"ConnectionCredentialPromptDialog.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Connection credential prompt source is readable for modal quit guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t showModalFunction = source.find("HRESULT ConnectionCredentialPromptWindow::ShowModal");
    const size_t onCreateFunction  = source.find("bool ConnectionCredentialPromptWindow::OnCreate", showModalFunction);
    Require(showModalFunction != std::string::npos && onCreateFunction != std::string::npos && showModalFunction < onCreateFunction,
            "Connection credential prompt ShowModal source block is found");

    const std::string showModalBlock = source.substr(showModalFunction, onCreateFunction - showModalFunction);
    const size_t quitBranch          = showModalBlock.find("if (getMessageResult == 0)");
    const size_t repostQuit          = showModalBlock.find("PostQuitMessage", quitBranch);
    const size_t quitBreak           = showModalBlock.find("break;", quitBranch);
    Require(quitBranch != std::string::npos && repostQuit != std::string::npos && quitBreak != std::string::npos && repostQuit < quitBreak,
            "Connection credential prompt modal loop has a WM_QUIT branch");

    const std::string quitBlock = showModalBlock.substr(quitBranch, quitBreak - quitBranch);
    const size_t destroyWindow  = quitBlock.find("DestroyWindow(");
    const size_t resetWindow    = quitBlock.find("_hWnd.reset()");
    Require((destroyWindow != std::string::npos && destroyWindow < repostQuit - quitBranch) ||
                (resetWindow != std::string::npos && resetWindow < repostQuit - quitBranch),
            "Connection credential prompt destroys its HWND before returning from a WM_QUIT modal-loop exit");
}

void TestConnectionCredentialPromptTeardownDoesNotWipeThroughRawTextFieldPointers()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"RedSalamander" / L"ConnectionCredentialPromptDialog.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Connection credential prompt source is readable for secure-clear teardown guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    Require(source.find("ControlTreeContains(") == std::string::npos,
            "Connection credential prompt must not walk retained controls through raw child pointers during teardown");
    Require(source.find("SecureClearLiveSecretField") == std::string::npos,
            "Connection credential prompt must not secure-clear retained TextField storage through a raw child pointer");
    Require(source.find("_secretField->SecureClear") == std::string::npos,
            "Connection credential prompt must not secure-clear the retained secret TextField through a raw child pointer");
    Require(source.find("_userField->SecureClear") == std::string::npos,
            "Connection credential prompt must not secure-clear the retained user TextField through a raw child pointer");

    const size_t destructorStart = source.find("ConnectionCredentialPromptWindow::~ConnectionCredentialPromptWindow");
    const size_t helperStart     = source.find("void ConnectionCredentialPromptWindow::ClearControlPointers", destructorStart);
    Require(destructorStart != std::string::npos && helperStart != std::string::npos && destructorStart < helperStart,
            "Connection credential prompt destructor block is found");
    const std::string destructorBlock = source.substr(destructorStart, helperStart - destructorStart);
    Require(destructorBlock.find("_secretField") == std::string::npos,
            "Connection credential prompt destructor must not dereference retained child pointers after WM_NCDESTROY");
    Require(destructorBlock.find("_userField") == std::string::npos,
            "Connection credential prompt destructor must not dereference retained child pointers after WM_NCDESTROY");
    Require(destructorBlock.find("SecureClearLiveSecretField") == std::string::npos,
            "Connection credential prompt destructor must not secure-clear the live retained TextField after WM_NCDESTROY");
    Require(destructorBlock.find("_dxHost.Detach") == std::string::npos,
            "Connection credential prompt destructor must not perform late retained-tree teardown");

    const size_t windowProcStart = source.find("LRESULT ConnectionCredentialPromptWindow::WindowProc");
    const size_t wmNcDestroy     = source.find("if (message == WM_NCDESTROY)", windowProcStart);
    const size_t wmNcReturn      = source.find("return 0;", wmNcDestroy);
    Require(windowProcStart != std::string::npos && wmNcDestroy != std::string::npos && wmNcReturn != std::string::npos,
            "Connection credential prompt WM_NCDESTROY block is found");
    const std::string wmNcDestroyBlock = source.substr(wmNcDestroy, wmNcReturn - wmNcDestroy);

    const size_t detach   = wmNcDestroyBlock.find("_dxHost.Detach()");
    const size_t clearRaw = wmNcDestroyBlock.find("ClearControlPointers()");
    Require(clearRaw != std::string::npos && detach != std::string::npos && detach < clearRaw,
            "Connection credential prompt nulls retained child pointers after detaching the retained tree");

    const size_t clearHelper = source.find("void ConnectionCredentialPromptWindow::ClearControlPointers");
    const size_t nextMethod  = source.find("\nvoid ConnectionCredentialPromptWindow::", clearHelper + 1u);
    Require(clearHelper != std::string::npos && nextMethod != std::string::npos && clearHelper < nextMethod,
            "Connection credential prompt clear-pointer helper is found");
    const std::string clearHelperBlock = source.substr(clearHelper, nextMethod - clearHelper);
    const size_t secretFieldClear      = clearHelperBlock.find("_secretField");
    const size_t secretNull            = clearHelperBlock.find("nullptr", secretFieldClear);
    const size_t userFieldClear        = clearHelperBlock.find("_userField");
    const size_t userNull              = clearHelperBlock.find("nullptr", userFieldClear);
    Require(secretFieldClear != std::string::npos && secretNull != std::string::npos,
            "Connection credential prompt clears the raw secret-field pointer after retained tree detach");
    Require(userFieldClear != std::string::npos && userNull != std::string::npos,
            "Connection credential prompt clears the raw user-field pointer after retained tree detach");
}

void TestConnectionCredentialPromptUiaPumpDoesNotDetachTimedOutWorkers()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"RedSalamander" / L"SelfTest" / L"Commands" / L"Commands.SelfTest.Connections.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Connection selftest source is readable for UIA worker lifetime guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t helperStart = source.find("template <typename Task> [[nodiscard]] auto RunUiaTaskWithMessagePump(std::wstring_view label");
    const size_t helperEnd   = source.find("template <typename Task> [[nodiscard]] auto RunUiaTaskWithMessagePump(Task&& task)", helperStart + 1u);
    Require(helperStart != std::string::npos && helperEnd != std::string::npos && helperStart < helperEnd,
            "Connection selftest UIA pump helper block is found");
    const std::string helperBlock = source.substr(helperStart, helperEnd - helperStart);

    Require(helperBlock.find("worker.detach()") == std::string::npos,
            "Connection selftest UIA pump must never detach a worker that can outlive prompt teardown");
    const size_t requestStop = helperBlock.find("worker.request_stop()");
    const size_t joinWorker  = helperBlock.find("worker.join()");
    Require(requestStop != std::string::npos && joinWorker != std::string::npos && requestStop < joinWorker,
            "Connection selftest UIA pump requests stop but rejoins the worker before returning");
    const size_t defaultReturn = helperBlock.find("return Result{}");
    Require(defaultReturn == std::string::npos || joinWorker < defaultReturn,
            "Connection selftest UIA pump returns a timeout sentinel only after the worker has exited");
}

void TestWindowHostPaintUsesWilBeginPaintRaii()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.WindowHost.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "WindowHost source is readable for paint RAII guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    Require(source.find("struct ScopedPaint") == std::string::npos, "WindowHost must not use a hand-rolled paint RAII wrapper");
    Require(source.find("ScopedPaint paint") == std::string::npos, "WindowHost WM_PAINT must not instantiate the hand-rolled paint wrapper");
    Require(source.find("EndPaint(") == std::string::npos, "WindowHost paint cleanup must be handled by wil::unique_hdc_paint");
    Require(source.find("wil::BeginPaint") != std::string::npos, "WindowHost WM_PAINT uses wil::BeginPaint for automatic EndPaint cleanup");
}

RedSalamander::DxUi::WindowHostBitmapCapture CaptureAttachedHostWindowBitmapForWindowHostSuite(AttachedHostWindow& window, const char* context)
{
    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

    RedSalamander::DxUi::WindowHostBitmapCapture capture;
    Require(window.Host().DebugCaptureBitmap(capture), context);
    return capture;
}

[[nodiscard]] uint32_t GetWindowHostCapturePixelBgra(const RedSalamander::DxUi::WindowHostBitmapCapture& capture, UINT xPx, UINT yPx) noexcept
{
    if (xPx >= capture.widthPx || yPx >= capture.heightPx)
    {
        return 0u;
    }

    const size_t base = (static_cast<size_t>(yPx) * static_cast<size_t>(capture.widthPx) + static_cast<size_t>(xPx)) * 4u;
    if ((base + 3u) >= capture.bgraPixels.size())
    {
        return 0u;
    }

    return static_cast<uint32_t>(capture.bgraPixels[base + 0u]) | (static_cast<uint32_t>(capture.bgraPixels[base + 1u]) << 8u) |
           (static_cast<uint32_t>(capture.bgraPixels[base + 2u]) << 16u) | (static_cast<uint32_t>(capture.bgraPixels[base + 3u]) << 24u);
}

[[nodiscard]] UINT DipToPixelForWindowHost(float dip, float dpi, UINT maxPx) noexcept
{
    if (maxPx == 0u)
    {
        return 0u;
    }

    return (std::min)(maxPx - 1u, static_cast<UINT>((std::max)(0l, std::lround(static_cast<double>(dip) * static_cast<double>(dpi) / 96.0))));
}

[[nodiscard]] std::filesystem::path GetWindowHostPerfJsonlPathFromEnvironment()
{
    const DWORD required = GetEnvironmentVariableW(L"REDSALAMANDER_PERF_JSONL_PATH", nullptr, 0u);
    if (required == 0u)
    {
        return {};
    }

    std::wstring value(required, L'\0');
    const DWORD copied = GetEnvironmentVariableW(L"REDSALAMANDER_PERF_JSONL_PATH", value.data(), required);
    if (copied == 0u)
    {
        return {};
    }

    value.resize(copied);
    return std::filesystem::path(value);
}

[[nodiscard]] uintmax_t GetWindowHostFileSizeOrZero(const std::filesystem::path& path) noexcept
{
    std::error_code ec;
    const uintmax_t size = std::filesystem::file_size(path, ec);
    return ec ? 0u : size;
}

[[nodiscard]] std::string ReadWindowHostPerfJsonlFromOffset(const std::filesystem::path& path, uintmax_t offset)
{
    std::ifstream input(path, std::ios::binary);
    if (! input)
    {
        return {};
    }

    input.seekg(static_cast<std::streamoff>(offset), std::ios::beg);
    return std::string(std::istreambuf_iterator<char>(input), std::istreambuf_iterator<char>());
}

class ScopedWindowHostPerfJsonl final
{
public:
    ScopedWindowHostPerfJsonl()
    {
        _path = GetWindowHostPerfJsonlPathFromEnvironment();
        if (! _path.empty())
        {
            return;
        }

        _path = GetDxUiTestArtifactPath(L"dxui_windowhost_stage_metrics_testlocal.jsonl");
        std::error_code ec;
        std::filesystem::remove(_path, ec);
        Debug::Perf::ConfigureJsonlOutput(_path, L"DxUiTests", L"Debug");
        _ownsConfiguration = true;
    }

    ~ScopedWindowHostPerfJsonl()
    {
        if (_ownsConfiguration)
        {
            Debug::Perf::ClearJsonlOutput();
        }
    }

    ScopedWindowHostPerfJsonl(const ScopedWindowHostPerfJsonl&)            = delete;
    ScopedWindowHostPerfJsonl& operator=(const ScopedWindowHostPerfJsonl&) = delete;
    ScopedWindowHostPerfJsonl(ScopedWindowHostPerfJsonl&&)                 = delete;
    ScopedWindowHostPerfJsonl& operator=(ScopedWindowHostPerfJsonl&&)      = delete;

    [[nodiscard]] const std::filesystem::path& Path() const noexcept
    {
        return _path;
    }

private:
    std::filesystem::path _path;
    bool _ownsConfiguration = false;
};

struct ReentrantKeyboardMutationState
{
    size_t keyDownCount = 0u;
    size_t charCount    = 0u;
};

class ReentrantKeyboardMutationControl final : public RedSalamander::DxUi::Control
{
public:
    explicit ReentrantKeyboardMutationControl(ReentrantKeyboardMutationState& state) : _state(&state)
    {
        SetFocusable(true);
    }

    void Paint(RedSalamander::DxUi::WindowHost& /*host*/) const override
    {
    }

    bool OnKeyDown(RedSalamander::DxUi::WindowHost& /*host*/, UINT /*virtualKey*/, UINT /*modifiers*/) override
    {
        ++_state->keyDownCount;
        SetEnabled(false);
        return true;
    }

    bool OnChar(RedSalamander::DxUi::WindowHost& /*host*/, wchar_t /*ch*/, UINT /*modifiers*/) override
    {
        ++_state->charCount;
        SetEnabled(false);
        return true;
    }

private:
    ReentrantKeyboardMutationState* _state = nullptr;
};

class FocusLossMutatesNextFocusControl final : public RedSalamander::DxUi::Control
{
public:
    explicit FocusLossMutatesNextFocusControl(RedSalamander::DxUi::Control*& controlToDisable) : _controlToDisable(&controlToDisable)
    {
        SetFocusable(true);
    }

    void Paint(RedSalamander::DxUi::WindowHost& /*host*/) const override
    {
    }

protected:
    void OnFocusChanged(RedSalamander::DxUi::WindowHost& host, bool focused) override
    {
        if (! focused && _controlToDisable && *_controlToDisable)
        {
            (*_controlToDisable)->SetEnabled(false);
        }
        Control::OnFocusChanged(host, focused);
    }

private:
    RedSalamander::DxUi::Control** _controlToDisable = nullptr;
};

struct PostedPayloadDrainStressWindowState final
{
    PostedPayloadDrainStressWindowState()                                                      = default;
    PostedPayloadDrainStressWindowState(const PostedPayloadDrainStressWindowState&)            = delete;
    PostedPayloadDrainStressWindowState(PostedPayloadDrainStressWindowState&&)                 = delete;
    PostedPayloadDrainStressWindowState& operator=(const PostedPayloadDrainStressWindowState&) = delete;
    PostedPayloadDrainStressWindowState& operator=(PostedPayloadDrainStressWindowState&&)      = delete;

    std::atomic<uint32_t> deliveredCount{0};
    std::atomic<uint32_t> drainedCount{0};
    std::atomic<uint32_t> staleTokenCount{0};
    std::atomic<uint32_t> staleTokenRejectionCount{0};
    std::atomic<uint64_t> drainDurationUs{0};
    std::atomic<bool> payloadQueuedBeforeDrain{false};
    std::atomic<bool> payloadQueuedAfterDrain{false};
};

struct PostedPayloadDrainStressPayload final
{
    std::atomic<uint32_t>* destroyedCount = nullptr;
    std::vector<uint32_t> values;

    ~PostedPayloadDrainStressPayload()
    {
        if (destroyedCount)
        {
            destroyedCount->fetch_add(1u, std::memory_order_acq_rel);
        }
    }
};

LRESULT CALLBACK PostedPayloadDrainStressWndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    auto* state = reinterpret_cast<PostedPayloadDrainStressWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    switch (message)
    {
        case WM_NCCREATE:
        {
            const auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
            state              = static_cast<PostedPayloadDrainStressWindowState*>(create ? create->lpCreateParams : nullptr);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
            InitPostedPayloadWindow(hwnd);
            return TRUE;
        }

        case WndMsg::kFolderViewEnumerateComplete:
        {
            auto payload = TakeMessagePayload<PostedPayloadDrainStressPayload>(lParam);
            if (state && payload)
            {
                state->deliveredCount.fetch_add(1u, std::memory_order_acq_rel);
            }
            return 0;
        }

        case WM_NCDESTROY:
        {
            MSG queuedMessage{};
            if (state)
            {
                state->payloadQueuedBeforeDrain.store(
                    PeekMessageW(&queuedMessage, hwnd, WndMsg::kFolderViewEnumerateComplete, WndMsg::kFolderViewEnumerateComplete, PM_NOREMOVE) != 0,
                    std::memory_order_release);
            }
            const auto drainStarted = std::chrono::steady_clock::now();
            const size_t drained    = DrainPostedPayloadsForWindow(hwnd);
            const auto drainDurationUs =
                static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - drainStarted).count());
            if (state)
            {
                state->drainedCount.store(static_cast<uint32_t>(drained), std::memory_order_release);
                state->drainDurationUs.store(drainDurationUs, std::memory_order_release);
                state->payloadQueuedAfterDrain.store(
                    PeekMessageW(&queuedMessage, hwnd, WndMsg::kFolderViewEnumerateComplete, WndMsg::kFolderViewEnumerateComplete, PM_NOREMOVE) != 0,
                    std::memory_order_release);

                uint32_t staleTokenCount          = 0u;
                uint32_t staleTokenRejectionCount = 0u;
                while (PeekMessageW(&queuedMessage, hwnd, WndMsg::kFolderViewEnumerateComplete, WndMsg::kFolderViewEnumerateComplete, PM_REMOVE) != 0)
                {
                    ++staleTokenCount;
                    if (! TakeMessagePayload<PostedPayloadDrainStressPayload>(queuedMessage.lParam))
                    {
                        ++staleTokenRejectionCount;
                    }
                }
                state->staleTokenCount.store(staleTokenCount, std::memory_order_release);
                state->staleTokenRejectionCount.store(staleTokenRejectionCount, std::memory_order_release);
            }
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            break;
        }
    }

    return DefWindowProcW(hwnd, message, wParam, lParam);
}

[[nodiscard]] wil::unique_hwnd CreatePostedPayloadDrainStressWindow(PostedPayloadDrainStressWindowState& state)
{
    constexpr wchar_t kClassName[] = L"RedSalamander.DxUiTests.PostedPayloadDrainStressWindow";

    HINSTANCE instance = GetModuleHandleW(nullptr);
    WNDCLASSW existing{};
    if (GetClassInfoW(instance, kClassName, &existing) == FALSE)
    {
        WNDCLASSW wc{};
        wc.lpfnWndProc   = PostedPayloadDrainStressWndProc;
        wc.hInstance     = instance;
        wc.lpszClassName = kClassName;
        Require(RegisterClassW(&wc) != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS, "posted-payload drain stress window class registers");
    }

    return wil::unique_hwnd(CreateWindowExW(0, kClassName, L"", 0, 0, 0, 0, 0, HWND_MESSAGE, nullptr, instance, &state));
}

class OverlayZOrderControl final : public RedSalamander::DxUi::Control
{
public:
    void Paint(RedSalamander::DxUi::WindowHost& host) const override
    {
        auto* const dc = host.GetDeviceContext();
        if (! dc)
        {
            return;
        }

        if (auto* const brush = host.GetSolidBrush(D2D1::ColorF(0.86f, 0.08f, 0.06f, 1.0f)))
        {
            dc->FillRectangle(host.GetClientBoundsDip(), brush);
        }
    }

    void PaintOverlay(RedSalamander::DxUi::WindowHost& host) const override
    {
        auto* const dc = host.GetDeviceContext();
        if (! dc)
        {
            return;
        }

        if (auto* const brush = host.GetSolidBrush(D2D1::ColorF(0.04f, 0.86f, 0.20f, 1.0f)))
        {
            dc->FillRectangle(D2D1::RectF(12.0f, 12.0f, 64.0f, 56.0f), brush);
        }
    }
};

class AnimationTickTraceControl final : public RedSalamander::DxUi::Control
{
public:
    void Paint(RedSalamander::DxUi::WindowHost& /*host*/) const override
    {
    }

    bool Tick(RedSalamander::DxUi::WindowHost& /*host*/, uint64_t nowTickMs) override
    {
        lastTickMs = nowTickMs;
        ++tickCount;
        return true;
    }

    uint64_t tickCount  = 0u;
    uint64_t lastTickMs = 0u;
};

class OverlayHitRecordingControl final : public RedSalamander::DxUi::Control
{
public:
    explicit OverlayHitRecordingControl(TrackingControlState& state) : _state(&state)
    {
    }

    void Paint(RedSalamander::DxUi::WindowHost& /*host*/) const override
    {
    }

    bool OnMouseDown(RedSalamander::DxUi::WindowHost& /*host*/, D2D1_POINT_2F /*point*/, bool rightButton, UINT /*modifiers*/) override
    {
        ++_state->mouseDownCount;
        return ! rightButton;
    }

protected:
    [[nodiscard]] RedSalamander::DxUi::Control* HitTestOverlay(D2D1_POINT_2F point) override
    {
        return Control::HitTest(point);
    }

    [[nodiscard]] const RedSalamander::DxUi::Control* HitTestOverlay(D2D1_POINT_2F point) const override
    {
        return Control::HitTest(point);
    }

private:
    TrackingControlState* _state = nullptr;
};

struct SelfCapturingControlState
{
    size_t mouseDownCount          = 0u;
    size_t mouseMoveWhileDownCount = 0u;
    size_t captureLostCount        = 0u;
    bool dragging                  = false;
};

class SelfCapturingControl final : public RedSalamander::DxUi::Control
{
public:
    explicit SelfCapturingControl(SelfCapturingControlState& state) : _state(&state)
    {
    }

    void Paint(RedSalamander::DxUi::WindowHost& /*host*/) const override
    {
    }

    bool OnMouseDown(RedSalamander::DxUi::WindowHost& host, D2D1_POINT_2F /*point*/, bool rightButton, UINT /*modifiers*/) override
    {
        if (rightButton)
        {
            return false;
        }

        ++_state->mouseDownCount;
        _state->dragging = true;
        host.CaptureMouse(this);
        return true;
    }

    bool OnMouseMove(RedSalamander::DxUi::WindowHost& /*host*/, D2D1_POINT_2F /*point*/, UINT /*modifiers*/) override
    {
        if (_state->dragging)
        {
            ++_state->mouseMoveWhileDownCount;
        }
        return true;
    }

    bool OnMouseUp(RedSalamander::DxUi::WindowHost& /*host*/, D2D1_POINT_2F /*point*/, bool /*rightButton*/, UINT /*modifiers*/) override
    {
        _state->dragging = false;
        return true;
    }

    void OnCaptureLost(RedSalamander::DxUi::WindowHost& /*host*/) override
    {
        ++_state->captureLostCount;
        _state->dragging = false;
    }

private:
    SelfCapturingControlState* _state = nullptr;
};

struct RootReplacingPointerControlState
{
    size_t mouseDownCount = 0u;
    size_t mouseUpCount   = 0u;
};

class RootReplacingPointerControl final : public RedSalamander::DxUi::Control
{
public:
    explicit RootReplacingPointerControl(RootReplacingPointerControlState& state) : _state(&state)
    {
        SetFocusable(true);
    }

    void Paint(RedSalamander::DxUi::WindowHost& /*host*/) const override
    {
    }

    bool OnMouseDown(RedSalamander::DxUi::WindowHost& host, D2D1_POINT_2F /*point*/, bool rightButton, UINT /*modifiers*/) override
    {
        if (rightButton)
        {
            return false;
        }

        ++_state->mouseDownCount;
        host.SetRoot(std::make_unique<RedSalamander::DxUi::Panel>());
        return true;
    }

    bool OnMouseUp(RedSalamander::DxUi::WindowHost& host, D2D1_POINT_2F /*point*/, bool rightButton, UINT /*modifiers*/) override
    {
        if (rightButton)
        {
            return false;
        }

        ++_state->mouseUpCount;
        host.SetRoot(std::make_unique<RedSalamander::DxUi::Panel>());
        return true;
    }

private:
    RootReplacingPointerControlState* _state = nullptr;
};

struct RootReplacingHoverControlState
{
    size_t hoverEnterCount = 0u;
    size_t mouseMoveCount  = 0u;
};

class RootReplacingHoverControl final : public RedSalamander::DxUi::Control
{
public:
    explicit RootReplacingHoverControl(RootReplacingHoverControlState& state) : _state(&state)
    {
    }

    void Paint(RedSalamander::DxUi::WindowHost& /*host*/) const override
    {
    }

    bool OnMouseMove(RedSalamander::DxUi::WindowHost& /*host*/, D2D1_POINT_2F /*point*/, UINT /*modifiers*/) override
    {
        ++_state->mouseMoveCount;
        return true;
    }

protected:
    void OnHoverChanged(RedSalamander::DxUi::WindowHost& host, bool hovered) override
    {
        Control::OnHoverChanged(host, hovered);

        if (hovered)
        {
            ++_state->hoverEnterCount;
            host.SetRoot(std::make_unique<RedSalamander::DxUi::Panel>());
        }
    }

private:
    RootReplacingHoverControlState* _state = nullptr;
};

class DetachOrderProbeControl final : public RedSalamander::DxUi::Control
{
public:
    DetachOrderProbeControl(DWORD ownerThreadId, uint32_t expectedAttachmentCount, bool& destroyed, bool& sawAttachmentAlive) noexcept
        : _ownerThreadId(ownerThreadId),
          _expectedAttachmentCount(expectedAttachmentCount),
          _destroyed(destroyed),
          _sawAttachmentAlive(sawAttachmentAlive)
    {
    }
    DetachOrderProbeControl(const DetachOrderProbeControl&)            = delete;
    DetachOrderProbeControl& operator=(const DetachOrderProbeControl&) = delete;
    DetachOrderProbeControl(DetachOrderProbeControl&&)                 = delete;
    DetachOrderProbeControl& operator=(DetachOrderProbeControl&&)      = delete;

    ~DetachOrderProbeControl() override
    {
        _destroyed          = true;
        _sawAttachmentAlive = RedSalamander::DxUi::DebugGetSharedWindowHostAttachmentCountForThread(_ownerThreadId) == _expectedAttachmentCount;
    }

    void Paint(RedSalamander::DxUi::WindowHost&) const override
    {
    }

private:
    DWORD _ownerThreadId              = 0u;
    uint32_t _expectedAttachmentCount = 0u;
    bool& _destroyed;
    bool& _sawAttachmentAlive;
};

class RenderLayoutMutationProbeControl final : public RedSalamander::DxUi::Control
{
public:
    void Paint(RedSalamander::DxUi::WindowHost& /*host*/) const override
    {
        if (_attemptedMutation)
        {
            return;
        }

        _attemptedMutation = true;
        const_cast<RenderLayoutMutationProbeControl*>(this)->SetBounds(D2D1::RectF(24.0f, 24.0f, 132.0f, 68.0f));
    }

private:
    mutable bool _attemptedMutation = false;
};

void TestWindowHostKeyboardInputMarksFocusVisible()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Require(host.GetInputModality() == InputModality::Pointer, "window host starts in pointer modality");
    Require(! host.IsKeyboardFocusVisible(), "window host starts with pointer-style focus visuals");

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, 'A', 0, handled));
    Require(handled, "keyboard input handled at host level");
    Require(host.GetInputModality() == InputModality::Keyboard, "keyboard input switches modality to keyboard");
    Require(host.IsKeyboardFocusVisible(), "keyboard input enables keyboard focus visuals");
}

void TestWindowHostPointerInputClearsKeyboardFocusVisible()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, 'A', 0, handled));
    Require(host.IsKeyboardFocusVisible(), "keyboard modality is active before pointer test");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, 0, handled));
    Require(handled, "pointer input handled at host level");
    Require(host.GetInputModality() == InputModality::Pointer, "pointer input switches modality back to pointer");
    Require(! host.IsKeyboardFocusVisible(), "pointer input clears keyboard-only focus visuals");
}

void TestWindowHostCrossThreadDetachReleasesOwnerThreadAttachmentCount()
{
    using namespace RedSalamander::DxUi;

    const DWORD ownerThreadId                   = GetCurrentThreadId();
    const size_t baselineAttachedHostCount      = DebugGetAttachedWindowHostCount();
    const uint32_t baselineOwnerAttachmentCount = DebugGetSharedWindowHostAttachmentCountForThread(ownerThreadId);

    AttachedHostWindow window;
    Require(DebugGetAttachedWindowHostCount() == (baselineAttachedHostCount + 1u), "attached host window registers one additional WindowHost");
    Require(DebugGetSharedWindowHostAttachmentCountForThread(ownerThreadId) == (baselineOwnerAttachmentCount + 1u),
            "attached host window increments the owner thread graphics attachment count");

    std::thread worker([&window] { window.Host().Detach(); });
    worker.join();

    Require(DebugGetAttachedWindowHostCount() == baselineAttachedHostCount, "cross-thread detach removes the host from the attached-host registry");
    Require(DebugGetSharedWindowHostAttachmentCountForThread(ownerThreadId) == baselineOwnerAttachmentCount,
            "cross-thread detach releases the original owner thread graphics attachment count");
}

void TestWindowHostDetachKeepsSharedGraphicsAttachmentUntilControlTreeDestroyed()
{
    using namespace RedSalamander::DxUi;

    const DWORD ownerThreadId                   = GetCurrentThreadId();
    const size_t baselineAttachedHostCount      = DebugGetAttachedWindowHostCount();
    const uint32_t baselineOwnerAttachmentCount = DebugGetSharedWindowHostAttachmentCountForThread(ownerThreadId);

    AttachedHostWindow window;
    const uint32_t attachedOwnerAttachmentCount = baselineOwnerAttachmentCount + 1u;
    bool probeDestroyed                         = false;
    bool probeSawAttachmentAlive                = false;
    window.Host().SetRoot(std::make_unique<DetachOrderProbeControl>(ownerThreadId, attachedOwnerAttachmentCount, probeDestroyed, probeSawAttachmentAlive));

    Require(DebugGetAttachedWindowHostCount() == (baselineAttachedHostCount + 1u), "attached host registers before detach-order validation");
    Require(DebugGetSharedWindowHostAttachmentCountForThread(ownerThreadId) == attachedOwnerAttachmentCount,
            "attached host increments graphics attachment count before detach-order validation");

    window.Host().Detach();

    Require(probeDestroyed, "detach destroys the retained control tree");
    Require(probeSawAttachmentAlive, "detach keeps the shared graphics attachment alive until the retained control tree is destroyed");
    Require(DebugGetAttachedWindowHostCount() == baselineAttachedHostCount, "detach removes host from the attached-host registry after order validation");
    Require(DebugGetSharedWindowHostAttachmentCountForThread(ownerThreadId) == baselineOwnerAttachmentCount,
            "detach releases graphics attachment count after retained host resources are destroyed");
}

void TestWindowHostDestructorDetachesBeforeMemberTeardown()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.WindowHost.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "WindowHost source is readable for destructor detach guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t destructorFunction = source.find("WindowHost::~WindowHost()");
    const size_t detachFunction     = source.find("void WindowHost::Detach()", destructorFunction);
    Require(destructorFunction != std::string::npos && detachFunction != std::string::npos && destructorFunction < detachFunction,
            "WindowHost destructor source block is found");

    const std::string destructorBlock = source.substr(destructorFunction, detachFunction - destructorFunction);
    Require(destructorBlock.find("Detach();") != std::string::npos,
            "WindowHost destructor must detach before native text input members and retained controls are torn down");
}

void TestWindowHostEmitsFrameStageMetricsForCaptureRender()
{
    using namespace RedSalamander::DxUi;

    ScopedWindowHostPerfJsonl perfJsonl;
    Require(! perfJsonl.Path().empty(), "window host stage metric test has a perf JSONL sink");

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* label = root->AddChild<Label>(L"Frame telemetry");
    label->SetBounds(D2D1::RectF(8.0f, 8.0f, 160.0f, 32.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().Invalidate();

    static_cast<void>(CaptureAttachedHostWindowBitmapForWindowHostSuite(window, "frame-stage metric warm-up capture succeeds"));

    const uintmax_t metricOffset = GetWindowHostFileSizeOrZero(perfJsonl.Path());
    WindowHostBitmapCapture capture;
    Require(window.Host().DebugCaptureBitmap(capture), "frame-stage metric direct debug capture render succeeds");
    Require(capture.widthPx > 0u && capture.heightPx > 0u && ! capture.bgraPixels.empty(), "frame-stage metric capture has pixels");

    const std::string appendedMetrics                          = ReadWindowHostPerfJsonlFromOffset(perfJsonl.Path(), metricOffset);
    constexpr std::array<std::string_view, 6> kExpectedMetrics = {{
        "\"metric\":\"dxui.frame.total_us\"",
        "\"metric\":\"dxui.frame.update_us\"",
        "\"metric\":\"dxui.frame.render_us\"",
        "\"metric\":\"dxui.frame.present_us\"",
        "\"metric\":\"dxui.frame.dirty_rect_count\"",
        "\"metric\":\"dxui.frame.dirty_rect_area_px\"",
    }};
    for (const std::string_view metric : kExpectedMetrics)
    {
        Require(appendedMetrics.find(metric) != std::string::npos, "window host capture render emits every frame-stage metric");
    }

    const auto findMetricLine = [&](std::string_view metric) noexcept -> std::string_view
    {
        const size_t metricPosition = appendedMetrics.find(metric);
        if (metricPosition == std::string::npos)
        {
            return {};
        }

        const size_t lineStart = appendedMetrics.rfind('\n', metricPosition);
        const size_t lineEnd   = appendedMetrics.find('\n', metricPosition);
        const size_t start     = lineStart == std::string::npos ? 0u : lineStart + 1u;
        const size_t end       = lineEnd == std::string::npos ? appendedMetrics.size() : lineEnd;
        return std::string_view(appendedMetrics).substr(start, end - start);
    };

    const std::string_view dirtyCountLine = findMetricLine("\"metric\":\"dxui.frame.dirty_rect_count\"");
    const std::string_view dirtyAreaLine  = findMetricLine("\"metric\":\"dxui.frame.dirty_rect_area_px\"");
    Require(dirtyCountLine.find("\"value\":0") != std::string_view::npos, "full-frame capture reports zero dirty rect count");
    Require(dirtyAreaLine.find("\"value\":0") != std::string_view::npos, "full-frame capture reports zero dirty rect area");
}

void TestWindowHostBlocksLayoutMutationDuringRender()
{
    using namespace RedSalamander::DxUi;

    ScopedWindowHostPerfJsonl perfJsonl;
    Require(! perfJsonl.Path().empty(), "window host render-mutation test has a perf JSONL sink");

    AttachedHostWindow window;
    auto root                       = std::make_unique<Panel>();
    auto* mutating                  = root->AddChild<RenderLayoutMutationProbeControl>();
    const D2D1_RECT_F initialBounds = D2D1::RectF(8.0f, 8.0f, 96.0f, 40.0f);
    mutating->SetBounds(initialBounds);
    window.Host().SetRoot(std::move(root));
    window.Host().Invalidate();

    const uintmax_t metricOffset = GetWindowHostFileSizeOrZero(perfJsonl.Path());
    static_cast<void>(CaptureAttachedHostWindowBitmapForWindowHostSuite(window, "render layout mutation capture succeeds"));

    const D2D1_RECT_F finalBounds = mutating->GetBounds();
    Require(finalBounds.left == initialBounds.left && finalBounds.top == initialBounds.top && finalBounds.right == initialBounds.right &&
                finalBounds.bottom == initialBounds.bottom,
            "render layout mutation keeps control bounds unchanged");

    const std::string appendedMetrics = ReadWindowHostPerfJsonlFromOffset(perfJsonl.Path(), metricOffset);
    Require(appendedMetrics.find("\"metric\":\"dxui.frame.render_layout_mutation_blocked\"") != std::string::npos,
            "render layout mutation emits blocked counter");
}

void TestPostMessagePayloadTeardownDrainDeletesUndeliveredPayloads()
{
    constexpr UINT kPayloadMessage = WndMsg::kFolderViewEnumerateComplete;
    PostedPayloadDrainStressWindowState state;
    std::atomic<uint32_t> destroyedCount{0};

    wil::unique_hwnd hwnd = CreatePostedPayloadDrainStressWindow(state);
    Require(hwnd != nullptr, "payload drain stress window is created");

    constexpr uint32_t kPayloadCount = 1024u;
    const auto postStarted           = std::chrono::steady_clock::now();
    for (uint32_t i = 0u; i < kPayloadCount; ++i)
    {
        auto payload            = std::make_unique<PostedPayloadDrainStressPayload>();
        payload->destroyedCount = &destroyedCount;
        Require(PostMessagePayload(hwnd.get(), kPayloadMessage, 0, std::move(payload)), "PostMessagePayload accepts payloads while the target window is alive");
    }
    const auto postDurationUs =
        static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - postStarted).count());
    Debug::Perf::Emit(L"dxui.posted_payload.post_batch_us", L"1024 queued payloads", postDurationUs, kPayloadCount, kPayloadCount, S_OK);

    MSG capturedStaleMessage{};
    Require(PeekMessageW(&capturedStaleMessage, nullptr, kPayloadMessage, kPayloadMessage, PM_NOREMOVE) != 0,
            "payload drain stress captures one queued message identity before teardown");
    const HWND retiredHwnd = hwnd.get();

    hwnd.reset();
    Debug::Perf::Emit(L"dxui.posted_payload.teardown_drain_us",
                      L"1024 queued payloads",
                      state.drainDurationUs.load(std::memory_order_acquire),
                      kPayloadCount,
                      kPayloadCount,
                      S_OK);

    Require(state.deliveredCount.load(std::memory_order_acquire) == 0u, "stress test destroys the window before delivery");
    Require(state.drainedCount.load(std::memory_order_acquire) == kPayloadCount, "WM_NCDESTROY drains all queued payloads");
    Require(destroyedCount.load(std::memory_order_acquire) == kPayloadCount, "drained payloads are deleted exactly once");
    Require(state.payloadQueuedBeforeDrain.load(std::memory_order_acquire),
            "the teardown test observes a payload message queued while the retiring HWND is still valid");
    Require(state.payloadQueuedAfterDrain.load(std::memory_order_acquire),
            "the opaque-token design deliberately leaves stale messages queued without retaining their storage");
    Require(state.staleTokenCount.load(std::memory_order_acquire) == kPayloadCount,
            "the teardown stress test pumps every deliberately retained stale token before Windows can discard it");
    Require(state.staleTokenRejectionCount.load(std::memory_order_acquire) == kPayloadCount,
            "every stale queued token is rejected after teardown invalidates the registry entries");

    InitPostedPayloadWindow(retiredHwnd);
    Require(destroyedCount.load(std::memory_order_acquire) == kPayloadCount, "pumping stale tokens after teardown cannot delete payload storage a second time");

    auto stalePayload = TakeMessagePayload<PostedPayloadDrainStressPayload>(capturedStaleMessage.lParam);
    Require(! stalePayload, "a stale queued lParam is rejected after its registered payload was drained");

    auto staleAfterSimulatedHwndReuse = TakeMessagePayload<PostedPayloadDrainStressPayload>(capturedStaleMessage.lParam);
    Require(! staleAfterSimulatedHwndReuse, "clearing the retired-HWND fence never makes a stale lParam ownable again");

    auto unregisteredPayload = TakeMessagePayload<PostedPayloadDrainStressPayload>(static_cast<LPARAM>(0x1234u));
    Require(! unregisteredPayload, "TakeMessagePayload never adopts an unregistered lParam");
}

void TestContiguousPostedPayloadCoalescingPreservesQueueOrderAndOperationKeys()
{
    constexpr UINT kPayloadMessage    = WM_APP + 0x370u;
    constexpr UINT kCompletionMessage = WM_APP + 0x371u;
    constexpr WPARAM kFirstOperation  = 41u;
    constexpr WPARAM kSecondOperation = 42u;

    PostedPayloadDrainStressWindowState state;
    std::atomic<uint32_t> destroyedCount{0};
    wil::unique_hwnd hwnd = CreatePostedPayloadDrainStressWindow(state);
    Require(hwnd != nullptr, "contiguous payload test window is created");

    const auto postPayload = [&](WPARAM operationKey, std::initializer_list<uint32_t> values)
    {
        auto payload            = std::make_unique<PostedPayloadDrainStressPayload>();
        payload->destroyedCount = &destroyedCount;
        payload->values.assign(values);
        Require(PostMessagePayload(hwnd.get(), kPayloadMessage, operationKey, std::move(payload)), "test payload posts successfully");
    };
    const auto removePayloadMessage = [&]()
    {
        MSG message{};
        Require(PeekMessageW(&message, hwnd.get(), kPayloadMessage, kPayloadMessage, PM_REMOVE) != 0, "expected payload message is queued");
        return message;
    };
    const auto takeQueuedPayload = [&]()
    {
        const MSG message = removePayloadMessage();
        return TakeMessagePayload<PostedPayloadDrainStressPayload>(message.lParam);
    };
    const auto appendValues = [](std::unique_ptr<PostedPayloadDrainStressPayload>& current, std::unique_ptr<PostedPayloadDrainStressPayload> newer) noexcept
    { current->values.insert(current->values.end(), newer->values.begin(), newer->values.end()); };

    postPayload(kFirstOperation, {1u});
    postPayload(kFirstOperation, {2u});
    postPayload(kSecondOperation, {100u});
    postPayload(kFirstOperation, {3u});
    MSG current             = removePayloadMessage();
    auto differentOperation = TakeAndCoalesceContiguousPostedPayloads<PostedPayloadDrainStressPayload>(
        hwnd.get(), kPayloadMessage, kFirstOperation, current.lParam, [](const PostedPayloadDrainStressPayload&, uint64_t) noexcept {
        return true;
    }, appendValues);
    Require(differentOperation.payload && differentOperation.payload->values == std::vector<uint32_t>{1u, 2u},
            "only contiguous payloads for the current operation are reduced");
    Require(differentOperation.drainedPayloadCount == 1u, "same-operation contiguous payload count is reported");
    Require(differentOperation.stoppedAtQueuedMessage && differentOperation.queuedMessage.wParam == kSecondOperation,
            "a different operation key remains at the queue head");
    differentOperation.payload.reset();
    takeQueuedPayload().reset();
    takeQueuedPayload().reset();

    postPayload(kFirstOperation, {4u});
    Require(PostMessageW(hwnd.get(), kCompletionMessage, kFirstOperation, 0) != 0, "completion marker posts successfully");
    postPayload(kFirstOperation, {5u});
    current                 = removePayloadMessage();
    auto completionBoundary = TakeAndCoalesceContiguousPostedPayloads<PostedPayloadDrainStressPayload>(
        hwnd.get(), kPayloadMessage, kFirstOperation, current.lParam, [](const PostedPayloadDrainStressPayload&, uint64_t) noexcept {
        return true;
    }, appendValues);
    Require(completionBoundary.payload && completionBoundary.payload->values == std::vector<uint32_t>{4u}, "completion behind progress is not bypassed");
    Require(completionBoundary.drainedPayloadCount == 0u && completionBoundary.stoppedAtQueuedMessage &&
                completionBoundary.queuedMessage.message == kCompletionMessage,
            "completion remains the next queue message");
    completionBoundary.payload.reset();
    MSG completion{};
    Require(PeekMessageW(&completion, nullptr, 0, 0, PM_REMOVE) != 0 && completion.message == kCompletionMessage,
            "completion marker retains its queue position");
    takeQueuedPayload().reset();

    postPayload(kFirstOperation, {6u});
    postPayload(kFirstOperation, {7u});
    postPayload(kFirstOperation, {8u});
    current       = removePayloadMessage();
    auto budgeted = TakeAndCoalesceContiguousPostedPayloads<PostedPayloadDrainStressPayload>(
        hwnd.get(), kPayloadMessage, kFirstOperation, current.lParam, [](const PostedPayloadDrainStressPayload&, uint64_t drainedPayloadCount) noexcept {
        return drainedPayloadCount < 1u;
    }, appendValues);
    Require(budgeted.payload && budgeted.payload->values == std::vector<uint32_t>{6u, 7u} && budgeted.drainedPayloadCount == 1u,
            "caller budget stops coalescing without losing the reduced payload");
    budgeted.payload.reset();
    auto finalPayload = takeQueuedPayload();
    Require(finalPayload && finalPayload->values == std::vector<uint32_t>{8u}, "the final snapshot remains queued after the budget boundary");
    finalPayload.reset();

    postPayload(kFirstOperation, {9u});
    postPayload(kFirstOperation, {10u});
    current        = removePayloadMessage();
    auto cancelled = TakeAndCoalesceContiguousPostedPayloads<PostedPayloadDrainStressPayload>(
        hwnd.get(), kPayloadMessage, kFirstOperation, current.lParam, [](const PostedPayloadDrainStressPayload&, uint64_t) noexcept {
        return false;
    }, appendValues);
    Require(cancelled.payload && cancelled.payload->values == std::vector<uint32_t>{9u} && cancelled.drainedPayloadCount == 0u,
            "caller cancellation predicate leaves later payloads untouched");
    cancelled.payload.reset();
    takeQueuedPayload().reset();

    hwnd.reset();
    Require(state.drainedCount.load(std::memory_order_acquire) == 0u, "all coalescing-test payload messages are explicitly consumed");
    Require(destroyedCount.load(std::memory_order_acquire) == 11u, "coalesced and queued payloads are each destroyed exactly once");
}

void TestSharedTestSupportPreservesSandboxAndEnvironmentPolicies()
{
    namespace TestSupport = RedSalamander::TestSupport;

    constexpr std::wstring_view kEnvironmentName{L"REDSALAMANDER_TEST_SUPPORT_SCOPE_CONTRACT"};
    static_cast<void>(SetEnvironmentVariableW(kEnvironmentName.data(), nullptr));
    {
        TestSupport::ScopedEnvironmentVariable scope(kEnvironmentName, L"temporary");
        const TestSupport::EnvironmentValue current = TestSupport::ReadEnvironmentValue(kEnvironmentName);
        Require(current.error == ERROR_SUCCESS && current.value && current.value.value() == L"temporary", "environment scope installs its temporary value");
    }
    const TestSupport::EnvironmentValue missing = TestSupport::ReadEnvironmentValue(kEnvironmentName);
    Require(missing.error == ERROR_SUCCESS && ! missing.value, "environment scope restores an originally missing value");

    Require(SetEnvironmentVariableW(kEnvironmentName.data(), L"original") != FALSE, "environment scope test installs an original value");
    {
        TestSupport::ScopedEnvironmentVariable scope(kEnvironmentName, std::nullopt);
        const TestSupport::EnvironmentValue current = TestSupport::ReadEnvironmentValue(kEnvironmentName);
        Require(current.error == ERROR_SUCCESS && ! current.value, "environment scope can temporarily remove a value");
    }
    const TestSupport::EnvironmentValue restored = TestSupport::ReadEnvironmentValue(kEnvironmentName);
    Require(restored.error == ERROR_SUCCESS && restored.value && restored.value.value() == L"original", "environment scope restores the exact original value");
    static_cast<void>(SetEnvironmentVariableW(kEnvironmentName.data(), nullptr));

    std::error_code ec;
    const std::filesystem::path outer =
        TestSupport::AcquireTestDirectory({.harnessSegment = L"dxui", .leafSegment = L"test-support-contract", .fallbackRunIdPrefix = L"dxui"}, ec);
    Require(! ec && ! outer.empty(), "test-support contract acquires an outer sandbox");

    {
        const std::wstring outerText = outer.wstring();
        TestSupport::ScopedEnvironmentVariable rootScope(TestSupport::kTestRootEnvironmentVariable, outerText);
        TestSupport::ScopedEnvironmentVariable runScope(TestSupport::kTestRunIdEnvironmentVariable, L"test-support-contract-run");

        const TestSupport::TestDirectoryOptions cleanOptions{
            .harnessSegment      = L"contract-harness",
            .leafSegment         = L"bad/leaf",
            .fallbackRunIdPrefix = L"unused",
        };
        const std::filesystem::path cleanDirectory = TestSupport::AcquireTestDirectory(cleanOptions, ec);
        Require(! ec && cleanDirectory.filename() == L"bad_leaf", "sandbox acquisition sanitizes unsafe leaf characters");
        std::filesystem::create_directories(cleanDirectory / L"sentinel", ec);
        Require(! ec, "sandbox clean-policy sentinel is created");
        const std::filesystem::path cleanedDirectory = TestSupport::AcquireTestDirectory(cleanOptions, ec);
        Require(! ec && cleanedDirectory == cleanDirectory && ! std::filesystem::exists(cleanDirectory / L"sentinel", ec),
                "clean-before-create removes prior case contents");

        TestSupport::TestDirectoryOptions retainedOptions = cleanOptions;
        retainedOptions.leafSegment                       = L"retained";
        retainedOptions.cleanExisting                     = false;
        const std::filesystem::path retainedDirectory     = TestSupport::AcquireTestDirectory(retainedOptions, ec);
        std::filesystem::create_directories(retainedDirectory / L"sentinel", ec);
        Require(! ec, "sandbox retain-policy sentinel is created");
        static_cast<void>(TestSupport::AcquireTestDirectory(retainedOptions, ec));
        Require(! ec && std::filesystem::exists(retainedDirectory / L"sentinel", ec), "no-clean acquisition preserves prior case contents");

        TestSupport::TestDirectoryOptions emptyLeafOptions = cleanOptions;
        emptyLeafOptions.leafSegment                       = L"";
        emptyLeafOptions.emptyLeafFallback                 = L"default";
        const std::filesystem::path emptyLeafDirectory     = TestSupport::AcquireTestDirectory(emptyLeafOptions, ec);
        Require(! ec && emptyLeafDirectory.filename() == L"default", "sandbox acquisition preserves the caller's empty-leaf fallback");

        const std::filesystem::path artifacts = TestSupport::AcquireTestDirectory({.harnessSegment      = L"contract-artifacts",
                                                                                   .fallbackRunIdPrefix = L"unused",
                                                                                   .kind                = TestSupport::TestDirectoryKind::Artifacts,
                                                                                   .includeLeafSegment  = false,
                                                                                   .cleanExisting       = false},
                                                                                  ec);
        Require(! ec && artifacts.filename() == L"contract-artifacts" && artifacts.parent_path().filename() == L"artifacts",
                "artifact acquisition omits the scratch leaf when requested");
    }

    std::filesystem::remove_all(outer, ec);
    Require(! ec, "test-support contract cleans its outer sandbox");
}

void TestSharedTestSupportPumpsMessagesAndBoundsSnapshotPolling()
{
    namespace TestSupport = RedSalamander::TestSupport;
    using namespace std::chrono_literals;

    wil::unique_hwnd messageWindow(CreateWindowExW(0u, L"STATIC", L"waiting", 0u, 0, 0, 0, 0, HWND_MESSAGE, nullptr, GetModuleHandleW(nullptr), nullptr));
    Require(messageWindow != nullptr, "message-pump contract creates a message-only window");
    constexpr UINT kTestMessage = WM_APP + 73u;
    Require(PostMessageW(messageWindow.get(), kTestMessage, 0u, 0) != FALSE, "message-pump contract queues a window message");

    const TestSupport::MessagePumpWaitResult pumped = TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        MSG pending{};
        return PeekMessageW(&pending, messageWindow.get(), kTestMessage, kTestMessage, PM_NOREMOVE) == FALSE;
    },
        {.timeout = 500ms, .pollInterval = 1ms, .operationName = L"message-only window update"});
    Require(pumped.conditionMet && pumped.dispatchedMessageCount >= 1u, "message-pump wait dispatches queued UI work instead of starving it");
    Require(pumped.timeoutDiagnostic.empty(), "successful message-pump waits do not report a timeout");

    const TestSupport::MessagePumpWaitResult timedOut =
        TestSupport::PumpMessagesUntil([]() noexcept { return false; }, {.timeout = 25ms, .pollInterval = 1ms, .operationName = L"bounded timeout contract"});
    Require(! timedOut.conditionMet && timedOut.elapsed >= 20ms && timedOut.elapsed < 500ms, "message-pump timeout stays bounded near its declared budget");
    Require(timedOut.timeoutDiagnostic.find(L"bounded timeout contract") != std::wstring::npos &&
                timedOut.timeoutDiagnostic.find(L"budget 25 ms") != std::wstring::npos,
            "message-pump timeout reports the operation and declared budget");

    struct Snapshot final
    {
        uint32_t sequence = 0u;
    } lastSnapshot{};
    uint32_t sequence = 0u;
    std::wstring timeoutDiagnostic;
    const bool snapshotReady = TestSupport::WaitForSnapshot<Snapshot>(
        [&](Snapshot& snapshot) noexcept
    {
        snapshot.sequence = ++sequence;
        return true;
    },
        [](const Snapshot& snapshot) noexcept { return snapshot.sequence >= 3u; },
        {.timeout = 250ms, .pollInterval = 1ms, .operationName = L"typed snapshot contract"},
        &lastSnapshot,
        &timeoutDiagnostic);
    Require(snapshotReady && lastSnapshot.sequence == 3u, "typed snapshot polling returns the matching snapshot");
    Require(timeoutDiagnostic.empty(), "successful typed snapshot polling leaves no timeout diagnostic");

    sequence                    = 0u;
    const bool snapshotTimedOut = TestSupport::WaitForSnapshot<Snapshot>(
        [&](Snapshot& snapshot) noexcept
    {
        snapshot.sequence = ++sequence;
        return true;
    },
        [](const Snapshot&) noexcept { return false; },
        {.timeout = 10ms, .pollInterval = 1ms, .operationName = L"typed snapshot timeout"},
        &lastSnapshot,
        &timeoutDiagnostic);
    Require(! snapshotTimedOut && lastSnapshot.sequence > 0u, "typed snapshot timeout preserves the most recently observed snapshot");
    Require(timeoutDiagnostic.find(L"typed snapshot timeout") != std::wstring::npos, "typed snapshot timeout reports its operation name");
}

void TestWindowHostMouseMoveUpdatesHoverTarget()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root = std::make_unique<Panel>();
    TrackingControlState firstState;
    TrackingControlState secondState;
    auto* first  = root->AddChild<TrackingControl>(firstState);
    auto* second = root->AddChild<TrackingControl>(secondState);
    root->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 80.0f));
    first->SetBounds(D2D1::RectF(0.0f, 0.0f, 100.0f, 28.0f));
    second->SetBounds(D2D1::RectF(110.0f, 0.0f, 210.0f, 28.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 80.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(12, 12), handled));
    Require(handled, "first hover move handled");
    Require(firstState.hoverEnterCount == 1u, "first control receives hover enter");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(128, 12), handled));
    Require(handled, "second hover move handled");
    Require(firstState.hoverLeaveCount == 1u, "first control receives hover leave when hover target changes");
    Require(secondState.hoverEnterCount == 1u, "second control receives hover enter when hover target changes");
}

[[nodiscard]] HWND CreateForeignPopupForWindowHostMouseLeaveTest(POINT screenPoint)
{
    static constexpr PCWSTR kClassName = L"RedSalamander.DxUiTests.ForeignPopup";
    static const ATOM atom             = []() noexcept
    {
        WNDCLASSW windowClass{};
        windowClass.lpfnWndProc   = DefWindowProcW;
        windowClass.hInstance     = GetModuleHandleW(nullptr);
        windowClass.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
        windowClass.lpszClassName = kClassName;
        return RegisterClassW(&windowClass);
    }();
    Require(atom != 0, "window host mouse-leave foreign popup class registers");

    return CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_TOPMOST,
                           kClassName,
                           L"DxUiTestsForeignPopup",
                           WS_POPUP | WS_VISIBLE,
                           screenPoint.x - 8,
                           screenPoint.y - 8,
                           48,
                           48,
                           nullptr,
                           nullptr,
                           GetModuleHandleW(nullptr),
                           nullptr);
}

void TestWindowHostMouseLeaveOverForeignPopupClearsHover()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    SetWindowPos(window.Hwnd(), nullptr, 180, 180, 320, 220, SWP_NOZORDER);
    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);

    auto root = std::make_unique<Panel>();
    TrackingControlState controlState;
    auto* control = root->AddChild<TrackingControl>(controlState);
    root->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 220.0f));
    control->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 80.0f));
    window.Host().SetRoot(std::move(root));
    window.PumpMessages();

    constexpr LONG kHoverX = 32;
    constexpr LONG kHoverY = 32;
    static_cast<void>(SendMessageW(window.Hwnd(), WM_MOUSEMOVE, 0, MAKELPARAM(kHoverX, kHoverY)));
    Require(controlState.hoverEnterCount == 1u, "window host mouse-leave popup test starts with a hovered control");

    POINT popupPoint{kHoverX, kHoverY};
    Require(ClientToScreen(window.Hwnd(), &popupPoint) != FALSE, "window host mouse-leave popup point converts to screen coordinates");
    wil::unique_hwnd popup(CreateForeignPopupForWindowHostMouseLeaveTest(popupPoint));
    Require(popup != nullptr, "window host mouse-leave popup creates a foreign popup over the host");
    SetWindowPos(popup.get(), HWND_TOPMOST, popupPoint.x - 8, popupPoint.y - 8, 48, 48, SWP_SHOWWINDOW | SWP_NOACTIVATE);
    UpdateWindow(popup.get());

    // The foreign popup is created over the host purely to document the scenario; the
    // production WM_MOUSELEAVE handler clears hover unconditionally (it reads neither the
    // live cursor nor WindowFromPoint), so the delivered WM_MOUSELEAVE alone exercises the
    // contract without warping the interactive user's cursor.
    static_cast<void>(SendMessageW(window.Hwnd(), WM_MOUSELEAVE, 0, 0));

    Require(controlState.hoverLeaveCount == 1u, "window host mouse leave over a foreign popup clears owner hover instead of rearming tracking");
}

void TestWindowHostMouseLeaveWithForeignCaptureClearsHover()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root = std::make_unique<Panel>();
    TrackingControlState controlState;
    auto* control = root->AddChild<TrackingControl>(controlState);
    root->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 220.0f));
    control->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 80.0f));
    window.Host().SetRoot(std::move(root));
    window.PumpMessages();

    constexpr LONG kHoverX = 32;
    constexpr LONG kHoverY = 32;
    static_cast<void>(SendMessageW(window.Hwnd(), WM_MOUSEMOVE, 0, MAKELPARAM(kHoverX, kHoverY)));
    Require(controlState.hoverEnterCount == 1u, "window host foreign-capture test starts with a hovered control");

    POINT popupPoint{260, 170};
    Require(ClientToScreen(window.Hwnd(), &popupPoint) != FALSE, "window host foreign-capture popup point converts to screen coordinates");
    wil::unique_hwnd popup(CreateForeignPopupForWindowHostMouseLeaveTest(popupPoint));
    Require(popup != nullptr, "window host foreign-capture test creates a foreign popup");
    SetWindowPos(popup.get(), HWND_TOPMOST, popupPoint.x - 8, popupPoint.y - 8, 48, 48, SWP_SHOWWINDOW | SWP_NOACTIVATE);
    UpdateWindow(popup.get());

    SetCapture(popup.get());
    const auto releaseCapture = wil::scope_exit([&]() noexcept
    {
        if (GetCapture() == popup.get())
        {
            ReleaseCapture();
        }
    });
    Require(GetCapture() == popup.get(), "window host foreign-capture test gives capture to a non-owner hwnd");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_MOUSELEAVE, 0, 0));

    Require(controlState.hoverLeaveCount == 1u, "window host mouse leave with foreign capture clears owner hover instead of rearming tracking");
}

void TestWindowHostTabTraversal()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* first  = root->AddChild<Button>(L"First");
    auto* second = root->AddChild<Button>(L"Second");
    auto* third  = root->AddChild<Button>(L"Third");
    first->SetBounds(D2D1::RectF(0.0f, 0.0f, 80.0f, 24.0f));
    second->SetBounds(D2D1::RectF(0.0f, 28.0f, 80.0f, 52.0f));
    third->SetBounds(D2D1::RectF(0.0f, 56.0f, 80.0f, 80.0f));

    host.SetRoot(std::move(root));
    host.SetFocusControl(first);
    Require(host.GetFocusControl() == first, "initial focus control");

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab key handled");
    Require(host.GetFocusControl() == second, "tab moves focus to second control");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "second tab key handled");
    Require(host.GetFocusControl() == third, "tab moves focus to third control");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "third tab key handled");
    Require(host.GetFocusControl() == first, "tab traversal wraps to the first control");
}

void TestWindowHostShiftTabTraversal()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* first  = root->AddChild<Button>(L"First");
    auto* second = root->AddChild<Button>(L"Second");
    auto* third  = root->AddChild<Button>(L"Third");
    first->SetBounds(D2D1::RectF(0.0f, 0.0f, 80.0f, 24.0f));
    second->SetBounds(D2D1::RectF(0.0f, 28.0f, 80.0f, 52.0f));
    third->SetBounds(D2D1::RectF(0.0f, 56.0f, 80.0f, 80.0f));

    host.SetRoot(std::move(root));
    host.SetFocusControl(first);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "shift+tab key handled");
    Require(host.GetFocusControl() == third, "shift+tab traversal wraps to the last control");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYUP, VK_SHIFT, 0, handled));
}

void TestWindowHostNativeFocusLossRetainsLogicalFocusForTraversal()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* first  = root->AddChild<Button>(L"First");
    auto* second = root->AddChild<Button>(L"Second");
    auto* third  = root->AddChild<Button>(L"Third");
    first->SetBounds(D2D1::RectF(0.0f, 0.0f, 80.0f, 24.0f));
    second->SetBounds(D2D1::RectF(0.0f, 28.0f, 80.0f, 52.0f));
    third->SetBounds(D2D1::RectF(0.0f, 56.0f, 80.0f, 80.0f));

    host.SetRoot(std::move(root));
    host.SetFocusControl(second);
    Require(host.GetFocusControl() == second, "native focus loss test starts from the second control");
    Require(second->HasFocus(), "native focus loss test starts with active focus visuals");

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KILLFOCUS, 0, 0, handled));
    Require(handled, "native focus loss is handled");
    Require(host.GetFocusControl() == second, "native focus loss keeps the retained logical focus target");
    Require(! second->HasFocus(), "native focus loss clears active control focus visuals");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SETFOCUS, 0, 0, handled));
    Require(handled, "native focus regain is handled");
    Require(host.GetFocusControl() == second, "native focus regain keeps the retained logical focus target");
    Require(second->HasFocus(), "native focus regain restores active control focus visuals");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));
    Require(handled, "shift key down is tracked before native focus loss");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KILLFOCUS, 0, 0, handled));
    Require(handled, "second native focus loss is handled");
    Require(host.GetFocusControl() == second, "second native focus loss still keeps the retained logical focus target");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab traversal after native focus loss is handled");
    Require(host.GetFocusControl() == third, "tab traversal after native focus loss continues from the retained control");
}

void TestWindowHostReturnInvokesDefaultButtonWhenFocusedControlDoesNotOwnEnter()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* field  = root->AddChild<TextField>();
    auto* button = root->AddChild<Button>(L"Search");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    button->SetBounds(D2D1::RectF(0.0f, 36.0f, 100.0f, 64.0f));

    size_t clickCount = 0u;
    button->SetOnClick([&clickCount] { ++clickCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(button);
    host.SetFocusControl(field);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled through host default-button routing");
    Require(clickCount == 1u, "default button invoked when focused field does not own enter");
}

void TestWindowHostReturnInvokesDefaultButtonWhenNoControlIsFocused()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"OK");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 100.0f, 28.0f));

    size_t clickCount = 0u;
    button->SetOnClick([&clickCount] { ++clickCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(button);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled through host default-button routing with no focused control");
    Require(clickCount == 1u, "default button invoked when no focused control owns enter");
}

void TestWindowHostReturnDoesNotInvokeDefaultButtonWhenFocusedControlOwnsEnter()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* field  = root->AddChild<TextField>();
    auto* button = root->AddChild<Button>(L"Search");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    button->SetBounds(D2D1::RectF(0.0f, 36.0f, 100.0f, 64.0f));

    bool submitted    = false;
    size_t clickCount = 0u;
    field->SetOnSubmitted([&submitted] { submitted = true; });
    button->SetOnClick([&clickCount] { ++clickCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(button);
    host.SetFocusControl(field);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled when focused field owns enter");
    Require(submitted, "focused field submit callback runs");
    Require(clickCount == 0u, "default button is not invoked when focused field owns enter");
}

void TestButtonKeyboardActivationCanReplaceRootSafely()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Apply");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 100.0f, 28.0f));

    bool clicked = false;
    button->SetOnClick([&]
    {
        clicked = true;
        host.SetRoot(std::make_unique<Panel>());
    });

    host.SetRoot(std::move(root));

    Require(button->OnKeyDown(host, VK_RETURN, 0), "button handles keyboard activation before replacing the root");
    Require(clicked, "button click callback runs before replacing the root");
    Require(host.GetRoot() != nullptr, "button keyboard activation can replace the root safely");
}

void TestWindowHostSpaceAndReturnInvokeFocusedButtonWithoutDefaultButtonFallback()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root           = std::make_unique<Panel>();
    auto* focusedButton = root->AddChild<Button>(L"Apply");
    auto* defaultButton = root->AddChild<Button>(L"Search");
    focusedButton->SetBounds(D2D1::RectF(0.0f, 0.0f, 100.0f, 28.0f));
    defaultButton->SetBounds(D2D1::RectF(0.0f, 36.0f, 100.0f, 64.0f));

    size_t focusedClickCount = 0u;
    size_t defaultClickCount = 0u;
    focusedButton->SetOnClick([&] { ++focusedClickCount; });
    defaultButton->SetOnClick([&] { ++defaultClickCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(defaultButton);
    host.SetFocusControl(focusedButton);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SPACE, 0, handled));
    Require(handled, "space handled by focused button");
    Require(focusedClickCount == 1u, "space invokes the focused button");
    Require(defaultClickCount == 0u, "space does not fall through to the default button");
    Require(host.GetInputModality() == InputModality::Keyboard, "space keeps keyboard input modality on the focused button");
    Require(host.IsKeyboardFocusVisible(), "space keeps keyboard focus visuals visible on the focused button");
    Require(host.GetFocusControl() == focusedButton, "space keeps focus on the focused button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled by focused button");
    Require(focusedClickCount == 2u, "return invokes the focused button");
    Require(defaultClickCount == 0u, "return does not fall through to the default button");
    Require(host.GetFocusControl() == focusedButton, "return keeps focus on the focused button");
}

void TestWindowHostDpiChangedIsHandled()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    RECT suggestedRect{10, 12, 210, 112};
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_DPICHANGED, MAKELONG(144u, 144u), reinterpret_cast<LPARAM>(&suggestedRect), handled));
    Require(handled, "window host handles WM_DPICHANGED explicitly");
}

void TestWindowHostDpiChangedInvalidatesMultilineCachesAndResizesAttachedWindow()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>();
    field->SetMultiline(true);
    field->SetText(L"alpha line\nbeta line\ncharlie line\ndelta line\necho line");
    field->SetBounds(D2D1::RectF(16.0f, 16.0f, 280.0f, 150.0f));
    window.Host().SetRoot(std::move(root));

    const UINT widthPxBefore  = window.Host().DebugGetWidthPx();
    const UINT heightPxBefore = window.Host().DebugGetHeightPx();
    static_cast<void>(CaptureAttachedHostWindowBitmapForWindowHostSuite(window, "attached host initial capture succeeds before dpi-change cache invalidation"));

    TextFieldDebugMultilineState beforeState{};
    Require(field->DebugGetMultilineState(window.Host(), beforeState), "multiline debug state available before dpi change");
    Require(beforeState.cachedLayoutPresent, "multiline text layout cache exists before dpi change");
    Require(! beforeState.layoutDirty, "multiline text layout cache is clean before dpi change");

    RECT windowRect{};
    Require(GetWindowRect(window.Hwnd(), &windowRect) != FALSE, "attached host window rect readable before dpi change");
    const LONG windowWidthPx  = std::max<LONG>(1, windowRect.right - windowRect.left);
    const LONG windowHeightPx = std::max<LONG>(1, windowRect.bottom - windowRect.top);
    RECT suggestedRect{windowRect.left,
                       windowRect.top,
                       windowRect.left + MulDiv(windowWidthPx, 144, USER_DEFAULT_SCREEN_DPI),
                       windowRect.top + MulDiv(windowHeightPx, 144, USER_DEFAULT_SCREEN_DPI)};

    const uint64_t invalidateCountBefore = window.Host().DebugGetInvalidateCount();
    bool handled                         = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_DPICHANGED, MAKELONG(144u, 144u), reinterpret_cast<LPARAM>(&suggestedRect), handled));
    Require(handled, "attached host handles WM_DPICHANGED");
    Require(std::abs(window.Host().GetDpi() - 144.0f) <= 0.001f, "attached host dpi updates to the requested value");
    Require(window.Host().DebugGetInvalidateCount() > invalidateCountBefore, "dpi change invalidates the attached host");

    TextFieldDebugMultilineState invalidatedState{};
    Require(field->DebugGetMultilineState(window.Host(), invalidatedState), "multiline debug state available immediately after dpi change");
    Require(! invalidatedState.cachedLayoutPresent, "dpi change clears the retained multiline text layout cache");
    Require(invalidatedState.layoutDirty, "dpi change marks multiline layout dirty before the next paint");

    window.PumpMessages();
    static_cast<void>(CaptureAttachedHostWindowBitmapForWindowHostSuite(window, "attached host capture succeeds after dpi-change cache invalidation"));

    TextFieldDebugMultilineState rebuiltState{};
    Require(field->DebugGetMultilineState(window.Host(), rebuiltState), "multiline debug state available after dpi-change repaint");
    Require(rebuiltState.cachedLayoutPresent, "multiline text layout cache is rebuilt on the first repaint after dpi change");
    Require(! rebuiltState.layoutDirty, "multiline text layout cache is clean again after repaint");
    Require(window.Host().DebugGetWidthPx() > widthPxBefore, "attached host client width grows after dpi-driven resize");
    Require(window.Host().DebugGetHeightPx() > heightPxBefore, "attached host client height grows after dpi-driven resize");
}

void TestWindowHostDpiChangedRecreatesTargetBitmapWhenSizeIsUnchanged()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.WindowHost.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "WindowHost source is readable for DPI target bitmap recreation guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t dpiChangedStart = source.find("void WindowHost::OnDpiChanged");
    const size_t handleStart     = source.find("LRESULT WindowHost::HandleMessage", dpiChangedStart);
    Require(dpiChangedStart != std::string::npos && handleStart != std::string::npos && dpiChangedStart < handleStart,
            "WindowHost OnDpiChanged source block is found");

    const std::string dpiChangedBlock = source.substr(dpiChangedStart, handleStart - dpiChangedStart);
    const size_t setDpi               = dpiChangedBlock.find("_d2dContext->SetDpi");
    const size_t recreateTarget       = dpiChangedBlock.find("PrepareForSwapChainResize()");
    const size_t invalidate           = dpiChangedBlock.rfind("Invalidate()");
    Require(setDpi != std::string::npos && recreateTarget != std::string::npos && invalidate != std::string::npos,
            "DPI change updates D2D DPI, detaches the old target bitmap, and invalidates");
    Require(setDpi < recreateTarget && recreateTarget < invalidate, "DPI change detaches the old target bitmap after applying the new DPI and before repaint");
}

void TestWindowHostAttachedWindowsRenderAcrossUiThreads()
{
    using namespace RedSalamander::DxUi;

    const auto installScene = [](AttachedHostWindow& window, std::wstring text)
    {
        auto root   = std::make_unique<Panel>();
        auto* label = root->AddChild<Label>(std::move(text));
        label->SetBounds(D2D1::RectF(16.0f, 16.0f, 220.0f, 48.0f));
        window.Host().SetRoot(std::move(root));
    };

    AttachedHostWindow primaryWindow;
    installScene(primaryWindow, L"Primary thread");
    const auto primaryCaptureBefore =
        CaptureAttachedHostWindowBitmapForWindowHostSuite(primaryWindow, "primary thread attached host capture succeeds before worker-thread render");
    Require(primaryCaptureBefore.widthPx > 0u && primaryCaptureBefore.heightPx > 0u,
            "primary thread capture has non-zero dimensions before worker-thread render");
    Require(primaryWindow.Host().DebugGetRenderCount() > 0u, "primary thread host renders before worker-thread render");

    struct ThreadRenderResult
    {
        bool captureSucceeded = false;
        UINT widthPx          = 0u;
        UINT heightPx         = 0u;
        uint64_t renderCount  = 0u;
    };

    std::promise<ThreadRenderResult> resultPromise;
    std::future<ThreadRenderResult> resultFuture = resultPromise.get_future();
    std::thread worker([promise = std::move(resultPromise), installScene]() mutable
    {
        ThreadRenderResult result{};
        AttachedHostWindow workerWindow;
        installScene(workerWindow, L"Worker thread");
        const auto workerCapture = CaptureAttachedHostWindowBitmapForWindowHostSuite(workerWindow, "worker thread attached host capture succeeds");
        result.captureSucceeded  = true;
        result.widthPx           = workerCapture.widthPx;
        result.heightPx          = workerCapture.heightPx;
        result.renderCount       = workerWindow.Host().DebugGetRenderCount();
        promise.set_value(result);
    });

    const ThreadRenderResult workerResult = resultFuture.get();
    worker.join();

    Require(workerResult.captureSucceeded, "worker thread attached host capture completed");
    Require(workerResult.widthPx > 0u && workerResult.heightPx > 0u, "worker thread capture has non-zero dimensions");
    Require(workerResult.renderCount > 0u, "worker thread host renders on its own UI thread");

    const auto primaryCaptureAfter =
        CaptureAttachedHostWindowBitmapForWindowHostSuite(primaryWindow, "primary thread attached host capture still succeeds after worker-thread render");
    Require(primaryCaptureAfter.widthPx > 0u && primaryCaptureAfter.heightPx > 0u, "primary thread capture has non-zero dimensions after worker-thread render");
    Require(primaryWindow.Host().DebugGetRenderCount() > 1u, "primary thread host renders again after worker-thread render");
}

void TestWindowHostEscapeInvokesCancelButton()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root          = std::make_unique<Panel>();
    auto* focusButton  = root->AddChild<Button>(L"Other");
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    focusButton->SetBounds(D2D1::RectF(0.0f, 0.0f, 100.0f, 28.0f));
    cancelButton->SetBounds(D2D1::RectF(0.0f, 36.0f, 100.0f, 64.0f));

    size_t cancelCount = 0u;
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    host.SetRoot(std::move(root));
    host.SetCancelButton(cancelButton);
    host.SetFocusControl(focusButton);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_ESCAPE, 0, handled));
    Require(handled, "escape handled through host cancel-button routing");
    Require(cancelCount == 1u, "cancel button invoked from escape");
}

void TestWindowHostEscapeClosesComboPopupBeforeCancelButton()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root          = std::make_unique<Panel>();
    auto* combo        = root->AddChild<ComboBox>();
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    cancelButton->SetBounds(D2D1::RectF(0.0f, 36.0f, 100.0f, 64.0f));
    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}});

    size_t cancelCount = 0u;
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    host.SetRoot(std::move(root));
    host.SetCancelButton(cancelButton);
    host.SetFocusControl(combo);

    Require(combo->OnKeyDown(host, VK_RETURN, 0), "combo enter opens popup before escape test");
    Require(combo->GetHitBounds().bottom > combo->GetBounds().bottom, "combo popup is open before escape");

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_ESCAPE, 0, handled));
    Require(handled, "escape handled while combo popup is open");
    Require(cancelCount == 0u, "cancel button not invoked while popup-owned escape is handled");
    Require(combo->GetHitBounds().bottom == combo->GetBounds().bottom, "escape closes combo popup first");
}

void TestWindowHostMenuKeyInvokesFocusedButtonContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Run");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 32.0f));

    RecordingContextMenuInvocation contextMenu;
    button->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 48.0f));
    host.SetFocusControl(button);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_APPS, 0, handled));
    Require(handled, "menu key handled for focused button");
    Require(contextMenu.count == 1u, "menu key invokes button context menu once");
    Require(contextMenu.lastKeyboardInvocation, "button menu key reports keyboard invocation");
    RequirePointNear(contextMenu.lastPoint, POINT{16, 16}, "button menu key uses the button keyboard anchor");
    Require(host.GetFocusControl() == button, "button menu key keeps focus on the button");
}

void TestWindowHostShiftF10InvokesFocusedToggleContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* toggle = root->AddChild<Toggle>(L"Menu bar");
    toggle->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 40.0f));
    toggle->SetChecked(true);

    RecordingContextMenuInvocation contextMenu;
    toggle->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 56.0f));
    host.SetFocusControl(toggle);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSKEYDOWN, VK_F10, 0, handled));
    Require(handled, "shift+f10 handled for focused toggle");
    Require(contextMenu.count == 1u, "shift+f10 invokes toggle context menu once");
    Require(contextMenu.lastKeyboardInvocation, "toggle shift+f10 reports keyboard invocation");
    RequirePointNear(contextMenu.lastPoint, POINT{16, 20}, "toggle shift+f10 uses the toggle keyboard anchor");
    Require(toggle->IsChecked(), "toggle shift+f10 does not change the checked state");
    Require(host.GetFocusControl() == toggle, "toggle shift+f10 keeps focus on the toggle");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYUP, VK_SHIFT, 0, handled));
}

void TestWindowHostMenuKeyInvokesFocusedCheckboxContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root      = std::make_unique<Panel>();
    auto* checkbox = root->AddChild<Checkbox>(L"Selected");
    checkbox->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    checkbox->SetChecked(true);

    RecordingContextMenuInvocation contextMenu;
    checkbox->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 48.0f));
    host.SetFocusControl(checkbox);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_APPS, 0, handled));
    Require(handled, "menu key handled for focused checkbox");
    Require(contextMenu.count == 1u, "menu key invokes checkbox context menu once");
    Require(contextMenu.lastKeyboardInvocation, "checkbox menu key reports keyboard invocation");
    RequirePointNear(contextMenu.lastPoint, POINT{16, 16}, "checkbox menu key uses the checkbox keyboard anchor");
    Require(checkbox->IsChecked(), "checkbox menu key does not change the checked state");
    Require(host.GetFocusControl() == checkbox, "checkbox menu key keeps focus on the checkbox");
}

void TestWindowHostSpaceAndReturnToggleFocusedToggleWithoutDefaultButtonFallback()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* toggle = root->AddChild<Toggle>(L"Ascending");
    auto* button = root->AddChild<Button>(L"Apply");
    toggle->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 40.0f));
    button->SetBounds(D2D1::RectF(0.0f, 48.0f, 100.0f, 76.0f));

    size_t toggleCount       = 0u;
    size_t defaultClickCount = 0u;
    toggle->SetOnToggled([&](bool) { ++toggleCount; });
    button->SetOnClick([&] { ++defaultClickCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(button);
    host.SetFocusControl(toggle);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SPACE, 0, handled));
    Require(handled, "space handled by focused toggle");
    Require(toggle->IsChecked(), "space toggles the focused toggle on");
    Require(toggleCount == 1u, "space fires the toggle callback once");
    Require(defaultClickCount == 0u, "space does not fall through to the default button");
    Require(host.GetInputModality() == InputModality::Keyboard, "space keeps keyboard input modality");
    Require(host.IsKeyboardFocusVisible(), "space keeps keyboard focus visuals visible");
    Require(host.GetFocusControl() == toggle, "space keeps focus on the toggle");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled by focused toggle");
    Require(! toggle->IsChecked(), "return toggles the focused toggle off");
    Require(toggleCount == 2u, "return fires the toggle callback once");
    Require(defaultClickCount == 0u, "return does not fall through to the default button");
    Require(host.GetFocusControl() == toggle, "return keeps focus on the toggle");
}

void TestWindowHostSpaceTogglesFocusedCheckboxAndReturnInvokesDefaultButton()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root      = std::make_unique<Panel>();
    auto* checkbox = root->AddChild<Checkbox>(L"Selected");
    auto* button   = root->AddChild<Button>(L"Apply");
    checkbox->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    button->SetBounds(D2D1::RectF(0.0f, 40.0f, 100.0f, 68.0f));

    size_t toggleCount       = 0u;
    size_t defaultClickCount = 0u;
    checkbox->SetOnToggled([&](bool) { ++toggleCount; });
    button->SetOnClick([&] { ++defaultClickCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(button);
    host.SetFocusControl(checkbox);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SPACE, 0, handled));
    Require(handled, "space handled by focused checkbox");
    Require(checkbox->IsChecked(), "space toggles the focused checkbox on");
    Require(toggleCount == 1u, "space fires the checkbox callback once");
    Require(defaultClickCount == 0u, "space does not fall through to the default button");
    Require(host.GetInputModality() == InputModality::Keyboard, "space keeps keyboard input modality for the checkbox");
    Require(host.IsKeyboardFocusVisible(), "space keeps keyboard focus visuals visible for the checkbox");
    Require(host.GetFocusControl() == checkbox, "space keeps focus on the checkbox");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled by focused checkbox through the host default button");
    Require(checkbox->IsChecked(), "return does not toggle the focused checkbox off");
    Require(toggleCount == 1u, "return does not fire the checkbox callback");
    Require(defaultClickCount == 1u, "return falls through to the default button for the checkbox");
    Require(host.GetFocusControl() == checkbox, "return keeps focus on the checkbox");
}

void TestWindowHostMixedDialogKeyboardFlowKeepsCommandsOnFocusedControls()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root           = std::make_unique<Panel>();
    auto* field         = root->AddChild<TextField>();
    auto* combo         = root->AddChild<ComboBox>();
    auto* checkbox      = root->AddChild<Checkbox>(L"Selected");
    auto* toggle        = root->AddChild<Toggle>(L"Ascending");
    auto* applyButton   = root->AddChild<Button>(L"Apply");
    auto* defaultButton = root->AddChild<Button>(L"Search");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo->SetBounds(D2D1::RectF(0.0f, 36.0f, 180.0f, 64.0f));
    checkbox->SetBounds(D2D1::RectF(0.0f, 72.0f, 220.0f, 104.0f));
    toggle->SetBounds(D2D1::RectF(0.0f, 112.0f, 220.0f, 152.0f));
    applyButton->SetBounds(D2D1::RectF(0.0f, 160.0f, 100.0f, 188.0f));
    defaultButton->SetBounds(D2D1::RectF(108.0f, 160.0f, 208.0f, 188.0f));
    combo->SetItems({ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}});

    size_t defaultClickCount   = 0u;
    size_t applyClickCount     = 0u;
    size_t checkboxToggleCount = 0u;
    size_t toggleCount         = 0u;
    defaultButton->SetOnClick([&] { ++defaultClickCount; });
    applyButton->SetOnClick([&] { ++applyClickCount; });
    checkbox->SetOnToggled([&](bool) { ++checkboxToggleCount; });
    toggle->SetOnToggled([&](bool) { ++toggleCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(defaultButton);
    host.SetFocusControl(field);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled from focused text field in mixed dialog flow");
    Require(defaultClickCount == 1u, "focused text field falls back to the default button in mixed dialog flow");
    Require(applyClickCount == 0u, "text field return does not invoke the non-default command button");
    Require(host.GetFocusControl() == field, "text field return keeps focus on the field");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from text field handled in mixed dialog flow");
    Require(host.GetFocusControl() == combo, "tab advances focus from text field to combo");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SPACE, 0, handled));
    Require(handled, "space handled by focused combo in mixed dialog flow");
    Require(combo->DebugIsPopupOpen(), "focused combo opens its popup in mixed dialog flow");
    Require(defaultClickCount == 1u, "combo space does not leak to the default button in mixed dialog flow");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled by focused combo in mixed dialog flow");
    Require(! combo->DebugIsPopupOpen(), "focused combo closes its popup in mixed dialog flow");
    Require(defaultClickCount == 1u, "combo return does not leak to the default button in mixed dialog flow");
    Require(host.GetFocusControl() == combo, "combo command routing keeps focus on the combo");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from combo handled in mixed dialog flow");
    Require(host.GetFocusControl() == checkbox, "tab advances focus from combo to checkbox");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SPACE, 0, handled));
    Require(handled, "space handled by focused checkbox in mixed dialog flow");
    Require(checkbox->IsChecked(), "focused checkbox toggles in mixed dialog flow");
    Require(checkboxToggleCount == 1u, "checkbox callback fires once in mixed dialog flow");
    Require(defaultClickCount == 1u, "checkbox space does not leak to the default button in mixed dialog flow");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled from focused checkbox in mixed dialog flow");
    Require(checkbox->IsChecked(), "focused checkbox return does not toggle in mixed dialog flow");
    Require(checkboxToggleCount == 1u, "checkbox return does not fire the checkbox callback in mixed dialog flow");
    Require(defaultClickCount == 2u, "focused checkbox return invokes the default button in mixed dialog flow");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from checkbox handled in mixed dialog flow");
    Require(host.GetFocusControl() == toggle, "tab advances focus from checkbox to toggle");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled by focused toggle in mixed dialog flow");
    Require(toggle->IsChecked(), "focused toggle switches on in mixed dialog flow");
    Require(toggleCount == 1u, "toggle callback fires once in mixed dialog flow");
    Require(defaultClickCount == 2u, "toggle return does not leak to the default button in mixed dialog flow");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from toggle handled in mixed dialog flow");
    Require(host.GetFocusControl() == applyButton, "tab advances focus from toggle to command button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled by focused command button in mixed dialog flow");
    Require(applyClickCount == 1u, "focused command button invokes itself in mixed dialog flow");
    Require(defaultClickCount == 2u, "focused command button return does not leak to the default button in mixed dialog flow");
    Require(host.GetInputModality() == InputModality::Keyboard, "mixed dialog flow stays in keyboard modality");
    Require(host.IsKeyboardFocusVisible(), "mixed dialog flow preserves keyboard focus visuals");
    Require(host.GetFocusControl() == applyButton, "focused command button keeps focus after invocation in mixed dialog flow");
}

void TestWindowHostMixedDialogMouseFlowKeepsCommandsOnHitControls()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root           = std::make_unique<Panel>();
    auto* field         = root->AddChild<TextField>();
    auto* combo         = root->AddChild<ComboBox>();
    auto* checkbox      = root->AddChild<Checkbox>(L"Selected");
    auto* toggle        = root->AddChild<Toggle>(L"Ascending");
    auto* applyButton   = root->AddChild<Button>(L"Apply");
    auto* defaultButton = root->AddChild<Button>(L"Search");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo->SetBounds(D2D1::RectF(0.0f, 36.0f, 180.0f, 64.0f));
    checkbox->SetBounds(D2D1::RectF(0.0f, 72.0f, 220.0f, 104.0f));
    toggle->SetBounds(D2D1::RectF(0.0f, 112.0f, 220.0f, 152.0f));
    applyButton->SetBounds(D2D1::RectF(0.0f, 160.0f, 100.0f, 188.0f));
    defaultButton->SetBounds(D2D1::RectF(108.0f, 160.0f, 208.0f, 188.0f));
    combo->SetItems({ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}});

    size_t defaultClickCount   = 0u;
    size_t applyClickCount     = 0u;
    size_t checkboxToggleCount = 0u;
    size_t toggleCount         = 0u;
    defaultButton->SetOnClick([&] { ++defaultClickCount; });
    applyButton->SetOnClick([&] { ++applyClickCount; });
    checkbox->SetOnToggled([&](bool) { ++checkboxToggleCount; });
    toggle->SetOnToggled([&](bool) { ++toggleCount; });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 204.0f));
    host.SetDefaultButton(defaultButton);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, 'A', 0, handled));
    Require(host.IsKeyboardFocusVisible(), "keyboard modality is active before mixed mouse-flow coverage");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(16, 14), handled));
    Require(handled, "text field click handles mouse-down in mixed mouse flow");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(16, 14), handled));
    Require(host.GetFocusControl() == field, "text field click moves focus to the text field");
    Require(host.GetInputModality() == InputModality::Pointer, "text field click switches modality back to pointer");
    Require(! host.IsKeyboardFocusVisible(), "text field click clears keyboard focus visuals in mixed mouse flow");
    Require(defaultClickCount == 0u, "text field click does not invoke the default button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(172, 50), handled));
    Require(handled, "combo click handles mouse-down in mixed mouse flow");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(172, 50), handled));
    Require(combo->DebugIsPopupOpen(), "combo click opens the popup in mixed mouse flow");
    Require(host.GetFocusControl() == combo, "combo click moves focus to the combo");
    Require(defaultClickCount == 0u, "combo click does not invoke the default button");

    const D2D1_RECT_F popupItemRect = combo->DebugGetPopupItemRect(1u, &host);
    const LONG popupItemX           = static_cast<LONG>((popupItemRect.left + popupItemRect.right) * 0.5f);
    const LONG popupItemY           = static_cast<LONG>((popupItemRect.top + popupItemRect.bottom) * 0.5f);
    handled                         = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(popupItemX, popupItemY), handled));
    Require(handled, "combo popup row click handles mouse-down in mixed mouse flow");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(popupItemX, popupItemY), handled));
    Require(! combo->DebugIsPopupOpen(), "combo popup row click closes the popup in mixed mouse flow");
    Require(combo->GetSelectedIndex().has_value() && combo->GetSelectedIndex().value() == 1u, "combo popup row click commits the hit item in mixed mouse flow");
    Require(defaultClickCount == 0u, "combo popup row click does not invoke the default button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(28, 88), handled));
    Require(handled, "checkbox click handles mouse-down in mixed mouse flow");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(28, 88), handled));
    Require(checkbox->IsChecked(), "checkbox click toggles the checkbox on in mixed mouse flow");
    Require(checkboxToggleCount == 1u, "checkbox click fires the checkbox callback once in mixed mouse flow");
    Require(host.GetFocusControl() == checkbox, "checkbox click moves focus to the checkbox");
    Require(defaultClickCount == 0u, "checkbox click does not invoke the default button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(32, 132), handled));
    Require(handled, "toggle click handles mouse-down in mixed mouse flow");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(32, 132), handled));
    Require(toggle->IsChecked(), "toggle click toggles the switch on in mixed mouse flow");
    Require(toggleCount == 1u, "toggle click fires the toggle callback once in mixed mouse flow");
    Require(host.GetFocusControl() == toggle, "toggle click moves focus to the toggle");
    Require(defaultClickCount == 0u, "toggle click does not invoke the default button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(48, 174), handled));
    Require(handled, "command button click handles mouse-down in mixed mouse flow");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(48, 174), handled));
    Require(applyClickCount == 1u, "command button click invokes only the hit command button in mixed mouse flow");
    Require(defaultClickCount == 0u, "command button click does not fall through to the default button in mixed mouse flow");
    Require(host.GetFocusControl() == applyButton, "command button click moves focus to the hit button");
}

void TestWindowHostMenuKeyInvokesFocusedTreeContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* tree = root->AddChild<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 120.0f));

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 10u, .text = L"General"},
        TreeItemData{.id = 20u, .text = L"Viewers"},
        TreeItemData{.id = 30u, .text = L"Themes"},
    });

    RecordingTreeDelegate delegate;
    tree->SetModel(&model);
    tree->SetDelegate(&delegate);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));
    tree->SetSelectedItemId(20u);
    host.SetFocusControl(tree);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_APPS, 0, handled));
    Require(handled, "menu key handled for focused tree");
    Require(delegate.contextMenuCount == 1u, "menu key invokes tree context menu once");
    Require(delegate.lastContextMenuItemId == 20u, "menu key targets the selected tree item");
    RequirePointNear(delegate.lastContextMenuPoint, POINT{16, 44}, "menu key uses a stable selected-item anchor");
}

void TestWindowHostShiftF10InvokesFocusedTreeContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* tree = root->AddChild<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 120.0f));

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 10u, .text = L"General"},
        TreeItemData{.id = 20u, .text = L"Viewers"},
        TreeItemData{.id = 30u, .text = L"Themes"},
    });

    RecordingTreeDelegate delegate;
    tree->SetModel(&model);
    tree->SetDelegate(&delegate);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));
    tree->SetSelectedItemId(30u);
    host.SetFocusControl(tree);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));
    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSKEYDOWN, VK_F10, 0, handled));
    Require(handled, "shift+f10 handled for focused tree");
    Require(delegate.contextMenuCount == 1u, "shift+f10 invokes tree context menu once");
    Require(delegate.lastContextMenuItemId == 30u, "shift+f10 targets the selected tree item");
    RequirePointNear(delegate.lastContextMenuPoint, POINT{16, 72}, "shift+f10 uses the selected-item keyboard anchor");
}

void TestWindowHostMenuKeyInvokesFocusedGridContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));

    MultiRowGridModel model(3u);
    RecordingGridDelegate delegate;
    grid->SetModel(&model);
    grid->SetDelegate(&delegate);

    host.SetRoot(std::move(root));
    grid->GetSelectionModel().SetSingle(model.GetStableRowId(1u));
    host.SetFocusControl(grid);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_APPS, 0, handled));
    Require(handled, "menu key handled for focused grid");
    Require(delegate.contextMenuCount == 1u, "menu key invokes grid context menu once");
    Require(delegate.lastContextMenuRow == 1u, "menu key targets the selected grid row");
    RequirePointNear(delegate.lastContextMenuPoint, POINT{16, 74}, "menu key uses a stable selected-row anchor");
}

void TestWindowHostShiftF10InvokesFocusedGridContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));

    MultiRowGridModel model(3u);
    RecordingGridDelegate delegate;
    grid->SetModel(&model);
    grid->SetDelegate(&delegate);

    host.SetRoot(std::move(root));
    grid->GetSelectionModel().SetSingle(model.GetStableRowId(2u));
    host.SetFocusControl(grid);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSKEYDOWN, VK_F10, 0, handled));
    Require(handled, "shift+f10 handled for focused grid");
    Require(delegate.contextMenuCount == 1u, "shift+f10 invokes grid context menu once");
    Require(delegate.lastContextMenuRow == 2u, "shift+f10 targets the selected grid row");
    RequirePointNear(delegate.lastContextMenuPoint, POINT{16, 102}, "shift+f10 uses the selected-row keyboard anchor");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYUP, VK_SHIFT, 0, handled));
}

void TestWindowHostMenuKeyInvokesFocusedTextFieldContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    RecordingContextMenuInvocation contextMenu;
    field->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 40.0f));
    host.SetFocusControl(field);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_APPS, 0, handled));
    Require(handled, "menu key handled for focused text field");
    Require(contextMenu.count == 1u, "menu key invokes text field context menu once");
    Require(contextMenu.lastKeyboardInvocation, "text field menu key reports keyboard invocation");
    RequirePointNear(contextMenu.lastPoint, POINT{16, 14}, "text field menu key uses the text field keyboard anchor");
    Require(host.GetFocusControl() == field, "text field menu key keeps focus on the text field");
}

void TestWindowHostMenuKeyInvokesFocusedComboContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo->SetItems({ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}});
    combo->SetSelectedIndex(0u);

    RecordingContextMenuInvocation contextMenu;
    combo->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 40.0f));
    host.SetFocusControl(combo);

    Require(! combo->DebugIsPopupOpen(), "non-editable combo popup starts closed for menu-key context menu");

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_APPS, 0, handled));
    Require(handled, "menu key handled for focused non-editable combo");
    Require(contextMenu.count == 1u, "menu key invokes non-editable combo context menu once");
    Require(contextMenu.lastKeyboardInvocation, "non-editable combo menu key reports keyboard invocation");
    RequirePointNear(contextMenu.lastPoint, POINT{16, 14}, "non-editable combo menu key uses the combo keyboard anchor");
    Require(host.GetFocusControl() == combo, "non-editable combo menu key keeps focus on the combo");
    Require(! combo->DebugIsPopupOpen(), "non-editable combo menu key does not open the popup");
}

void TestWindowHostSetRootClearsDestroyedTreeInteractionState()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TrackingControlState oldState;
    auto oldRoot     = std::make_unique<Panel>();
    auto* oldControl = oldRoot->AddChild<TrackingControl>(oldState);
    oldControl->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 40.0f));

    host.SetRoot(std::move(oldRoot));
    host.SetFocusControl(oldControl);
    host.CaptureMouse(oldControl);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(8, 8), handled));
    Require(handled, "initial hover update handled before root swap");
    Require(oldState.focusGainCount == 1u, "old control receives focus before root swap");
    Require(oldState.hoverEnterCount == 1u, "old control receives hover before root swap");

    auto replacementRoot = std::make_unique<Panel>();
    host.SetRoot(std::move(replacementRoot));
    Require(host.GetFocusControl() == nullptr, "set root clears focused control from destroyed tree");
    Require(oldState.focusLossCount == 1u, "set root notifies old focused control about focus loss");
    Require(oldState.hoverLeaveCount == 1u, "set root clears hovered control from destroyed tree");
    Require(oldState.mouseLeaveCount == 1u, "set root issues a mouse-leave to the old hovered control");

    const size_t mouseMoveCountBefore = oldState.mouseMoveCount;
    const size_t mouseUpCountBefore   = oldState.mouseUpCount;

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(12, 12), handled));
    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(12, 12), handled));
    Require(oldState.mouseMoveCount == mouseMoveCountBefore, "stale hovered control is not reused after root swap");
    Require(oldState.mouseUpCount == mouseUpCountBefore, "stale captured control is not reused after root swap");
}

void TestWindowHostDetachDeactivatesSecureTextInputBeforeDestroyingRoot()
{
    using namespace RedSalamander::DxUi;

    class DestructionTrackedTextField final : public TextField
    {
    public:
        DestructionTrackedTextField(std::wstring text, size_t& destroyedCount) : TextField(std::move(text)), _destroyedCount(&destroyedCount)
        {
        }

        ~DestructionTrackedTextField() noexcept override
        {
            ++*_destroyedCount;
        }

    private:
        size_t* _destroyedCount = nullptr;
    };

    WindowHost host;
    size_t destroyedCount = 0u;
    auto root             = std::make_unique<Panel>();
    auto* secretField     = root->AddChild<DestructionTrackedTextField>(std::wstring(L"credential-secret"), destroyedCount);
    secretField->SetMasked(true);
    secretField->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));

    host.SetRoot(std::move(root));
    host.SetFocusControl(secretField);

    NativeTextInputState activeState{};
    Require(host.GetFocusControl() == secretField, "secure text field starts focused before detach");
    Require(host.HasActiveTextInput(), "secure text field has an active text-input session before detach");
    Require(host.DebugHasActiveNativeTextInputSession(), "secure text field has an active native text-input session before detach");
    Require(host.DebugGetNativeTextInputState(activeState), "secure text field exports native text-input state before detach");
    Require(activeState.text == L"credential-secret", "secure text field native cache contains the secret before detach");

    host.Detach();

    NativeTextInputState detachedState{};
    Require(destroyedCount == 1u, "detach destroys the secure text field through retained root ownership");
    Require(host.GetFocusControl() == nullptr, "detach clears focus before retained secure text field storage is destroyed");
    Require(! host.HasActiveTextInput(), "detach clears text-input session before retained secure text field storage is destroyed");
    Require(! host.DebugHasActiveNativeTextInputSession(), "detach clears native text-input session before retained secure text field storage is destroyed");
    Require(! host.DebugGetNativeTextInputState(detachedState), "detach clears the native text-input cache for destroyed secure fields");
}

void TestWindowHostClearChildrenPrunesDestroyedTreeInteractionState()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TrackingControlState oldState;
    auto root        = std::make_unique<Panel>();
    auto* rootPanel  = root.get();
    auto* oldControl = root->AddChild<TrackingControl>(oldState);
    oldControl->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 40.0f));

    host.SetRoot(std::move(root));
    host.SetFocusControl(oldControl);
    host.CaptureMouse(oldControl);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(8, 8), handled));
    Require(handled, "initial hover update handled before child clear");
    Require(host.GetFocusControl() == oldControl, "focus starts on the installed child control");

    rootPanel->ClearChildren();

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(12, 12), handled));
    Require(host.GetFocusControl() == nullptr, "clearing children prunes focus from destroyed controls on the next message");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(12, 12), handled));
    Require(oldState.mouseUpCount == 0u, "clearing children does not reuse stale captured controls after prune");
}

void TestWindowHostKeyDownCallbackDisablingFocusPrunesBeforePostDispatchSync()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root = std::make_unique<Panel>();
    ReentrantKeyboardMutationState state;
    auto* control = root->AddChild<ReentrantKeyboardMutationControl>(state);
    control->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 40.0f));

    host.SetRoot(std::move(root));
    host.SetFocusControl(control);
    Require(host.GetFocusControl() == control, "reentrant key-down test starts with the mutating control focused");

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, 'A', 0, handled));
    Require(handled, "focused mutating control handles key down");
    Require(state.keyDownCount == 1u, "focused mutating control receives one key-down callback");
    Require(host.GetFocusControl() == nullptr, "key-down callback disabling the focus target prunes retained focus before post-dispatch sync");
}

void TestWindowHostCharCallbackDisablingFocusPrunesBeforePostDispatchSync()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root = std::make_unique<Panel>();
    ReentrantKeyboardMutationState state;
    auto* control = root->AddChild<ReentrantKeyboardMutationControl>(state);
    control->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 40.0f));

    host.SetRoot(std::move(root));
    host.SetFocusControl(control);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_CHAR, L'x', 0, handled));
    Require(handled, "focused mutating control handles char input");
    Require(state.charCount == 1u, "focused mutating control receives one char callback");
    Require(host.GetFocusControl() == nullptr, "char callback disabling the focus target prunes retained focus before post-dispatch sync");
}

void TestWindowHostFocusLossCallbackRevalidatesRequestedFocusTarget()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root                 = std::make_unique<Panel>();
    Control* nextFocusControl = nullptr;
    auto* oldFocus            = root->AddChild<FocusLossMutatesNextFocusControl>(nextFocusControl);
    auto* nextFocus           = root->AddChild<Button>(L"Next");
    nextFocusControl          = nextFocus;
    oldFocus->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 32.0f));
    nextFocus->SetBounds(D2D1::RectF(0.0f, 40.0f, 120.0f, 72.0f));

    host.SetRoot(std::move(root));
    host.SetFocusControl(oldFocus);
    Require(host.GetFocusControl() == oldFocus, "focus-loss mutation test starts with old focus");

    host.SetFocusControl(nextFocus);
    Require(! nextFocus->IsEnabled(), "old focus loss callback disables the requested next focus control");
    Require(host.GetFocusControl() == nullptr, "focus change revalidates the requested focus target after focus-loss callbacks");
}

void TestWindowHostIgnoresObserverButtonsOutsideInstalledRoot()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root           = std::make_unique<Panel>();
    auto* field         = root->AddChild<TextField>();
    auto* defaultButton = root->AddChild<Button>(L"Search");
    auto* cancelButton  = root->AddChild<Button>(L"Cancel");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    defaultButton->SetBounds(D2D1::RectF(0.0f, 36.0f, 100.0f, 64.0f));
    cancelButton->SetBounds(D2D1::RectF(108.0f, 36.0f, 208.0f, 64.0f));

    Button outsideDefault(L"Outside Default");
    Button outsideCancel(L"Outside Cancel");

    size_t defaultCount        = 0u;
    size_t cancelCount         = 0u;
    size_t outsideDefaultCount = 0u;
    size_t outsideCancelCount  = 0u;
    defaultButton->SetOnClick([&defaultCount] { ++defaultCount; });
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });
    outsideDefault.SetOnClick([&outsideDefaultCount] { ++outsideDefaultCount; });
    outsideCancel.SetOnClick([&outsideCancelCount] { ++outsideCancelCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(defaultButton);
    host.SetCancelButton(cancelButton);
    host.SetDefaultButton(&outsideDefault);
    host.SetCancelButton(&outsideCancel);
    host.SetFocusControl(field);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return remains routed to the installed-root default button");
    Require(defaultCount == 1u, "outside default button does not replace the valid default button");
    Require(outsideDefaultCount == 0u, "outside default button is ignored");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_ESCAPE, 0, handled));
    Require(handled, "escape remains routed to the installed-root cancel button");
    Require(cancelCount == 1u, "outside cancel button does not replace the valid cancel button");
    Require(outsideCancelCount == 0u, "outside cancel button is ignored");
}

void TestWindowHostIgnoresFocusAndCaptureOutsideInstalledRoot()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TrackingControlState insideState;
    TrackingControlState outsideState;

    auto root           = std::make_unique<Panel>();
    auto* insideControl = root->AddChild<TrackingControl>(insideState);
    insideControl->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 40.0f));

    TrackingControl outsideControl(outsideState);
    outsideControl.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 40.0f));

    host.SetRoot(std::move(root));
    host.SetFocusControl(insideControl);
    Require(host.GetFocusControl() == insideControl, "focus starts on an installed-root control");
    Require(insideState.focusGainCount == 1u, "installed-root control receives focus");

    host.SetFocusControl(&outsideControl);
    Require(host.GetFocusControl() == insideControl, "outside control does not replace focused installed-root control");
    Require(outsideState.focusGainCount == 0u, "outside control does not receive focus");
    Require(insideState.focusLossCount == 0u, "installed-root focus is preserved when outside control is ignored");

    host.CaptureMouse(insideControl);
    host.CaptureMouse(&outsideControl);
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(8, 8), handled));
    Require(handled, "mouse move remains handled by the installed-root captured control");
    Require(insideState.mouseMoveCount == 1u, "installed-root captured control handles mouse move");
    Require(outsideState.mouseMoveCount == 0u, "outside captured control is ignored");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(8, 8), handled));
    Require(outsideState.mouseUpCount == 0u, "outside captured control does not receive mouse-up");
}

void TestWindowHostCaptureLossClearsPressedButtonState()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<ExposedButton>(L"Apply");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 32.0f));
    window.Host().SetRoot(std::move(root));

    bool handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(24, 16), handled));
    Require(handled, "capture-loss button test handles mouse-down");
    Require(button->IsPressed(), "capture-loss button test enters pressed state after mouse-down");

    handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_CAPTURECHANGED, 0, 0, handled));
    Require(handled, "capture-loss button test handles capture change");
    Require(! button->IsPressed(), "capture-loss button test clears pressed state when capture is lost");
}

void TestWindowHostRedundantCaptureDoesNotCancelMouseDownCapture()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    SelfCapturingControlState state;
    auto root     = std::make_unique<Panel>();
    auto* control = root->AddChild<SelfCapturingControl>(state);
    control->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 80.0f));
    window.Host().SetRoot(std::move(root));

    bool handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(24, 16), handled));
    Require(handled, "self-capturing control mouse-down is handled");
    Require(state.mouseDownCount == 1u, "self-capturing control receives one mouse-down");
    Require(state.captureLostCount == 0u, "self-capturing control keeps capture after WindowHost post-handler capture");
    Require(state.dragging, "self-capturing control remains in dragging state after mouse-down");

    handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_MOUSEMOVE, MK_LBUTTON, MAKELPARAM(24, 56), handled));
    Require(handled, "self-capturing control captured mouse-move is handled");
    Require(state.mouseMoveWhileDownCount == 1u, "self-capturing control receives captured mouse-move while dragging");

    handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_LBUTTONUP, 0, MAKELPARAM(24, 56), handled));
    Require(handled, "self-capturing control mouse-up is handled");
    Require(! state.dragging, "self-capturing control clears dragging state on mouse-up");
}

void TestWindowHostPointerDispatchDoesNotReuseTargetAfterRootReplacement()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    RootReplacingPointerControlState state;
    auto root     = std::make_unique<Panel>();
    auto* control = root->AddChild<RootReplacingPointerControl>(state);
    control->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 80.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 120.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(24, 16), handled));
    Require(handled, "root-replacing pointer-down is handled");
    Require(state.mouseDownCount == 1u, "root-replacing control receives one mouse-down");
    Require(host.GetFocusControl() == nullptr, "root-replacing pointer-down leaves no stale focus target");

    auto secondRoot     = std::make_unique<Panel>();
    auto* secondControl = secondRoot->AddChild<RootReplacingPointerControl>(state);
    secondControl->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 80.0f));
    host.SetRoot(std::move(secondRoot));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 120.0f));

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(24, 16), handled));
    Require(handled, "root-replacing pointer-up is handled");
    Require(state.mouseUpCount == 1u, "root-replacing control receives one mouse-up");

    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.WindowHost.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "WindowHost source is readable for reentrancy guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t downDispatch = source.find("const bool controlHandled");
    Require(downDispatch != std::string::npos, "WindowHost pointer-down dispatch marker exists");
    const size_t downPostDispatch = source.find("if (IsContextMenuDiagnosticsEnabled())", downDispatch);
    const size_t downReturn       = source.find("return 0;", downPostDispatch);
    Require(downPostDispatch != std::string::npos && downReturn != std::string::npos && downPostDispatch < downReturn,
            "WindowHost pointer-down post-dispatch block is found");
    const std::string downPostCallbackBlock = source.substr(downPostDispatch, downReturn - downPostDispatch);
    Require(downPostCallbackBlock.find("target->") == std::string::npos,
            "WindowHost pointer-down does not dereference target after OnMouseDown/OnMouseDoubleClick");
    Require(downPostCallbackBlock.find("DescribeWindowHostTraceControl(target)") == std::string::npos,
            "WindowHost pointer-down diagnostics do not describe target after OnMouseDown/OnMouseDoubleClick");

    const size_t upDispatch = source.find("target->OnMouseUp");
    Require(upDispatch != std::string::npos, "WindowHost pointer-up dispatch marker exists");
    const size_t upPostDispatch = source.find("if (IsContextMenuDiagnosticsEnabled())", upDispatch);
    const size_t upRelease      = source.find("ReleaseMouseCapture();", upPostDispatch);
    Require(upPostDispatch != std::string::npos && upRelease != std::string::npos && upPostDispatch < upRelease,
            "WindowHost pointer-up post-dispatch block is found");
    const std::string upPostCallbackBlock = source.substr(upPostDispatch, upRelease - upPostDispatch);
    Require(upPostCallbackBlock.find("DescribeWindowHostTraceControl(target)") == std::string::npos,
            "WindowHost pointer-up diagnostics do not describe target after OnMouseUp");
}

void TestWindowHostHoverEnterDoesNotReuseTargetAfterRootReplacement()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    RootReplacingHoverControlState state;
    auto root     = std::make_unique<Panel>();
    auto* control = root->AddChild<RootReplacingHoverControl>(state);
    control->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 80.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 120.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(24, 16), handled));
    Require(handled, "root-replacing hover-enter mouse move is handled");
    Require(state.hoverEnterCount == 1u, "root-replacing hover control receives hover enter");
    Require(state.mouseMoveCount == 0u, "root-replacing hover control is not reused for mouse move after replacing the root");

    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.WindowHost.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "WindowHost source is readable for hover reentrancy guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());
    const size_t updateHover = source.find("void WindowHost::UpdateHover");
    Require(updateHover != std::string::npos, "WindowHost UpdateHover source exists");
    const size_t hoverEnter = source.find("OnHoverChanged(*this, true)", updateHover);
    const size_t mouseMove  = source.find("OnMouseMove(*this, pointDip, modifiers)", hoverEnter);
    Require(hoverEnter != std::string::npos && mouseMove != std::string::npos && hoverEnter < mouseMove,
            "WindowHost UpdateHover hover-enter and mouse-move calls are found");
    const std::string hoverPostCallbackBlock = source.substr(hoverEnter, mouseMove - hoverEnter);
    Require(hoverPostCallbackBlock.find("RevalidateInteractiveDispatchedControl") != std::string::npos,
            "WindowHost UpdateHover revalidates the hover target after hover-enter callbacks");
}

void TestWindowHostDeviceLossSchedulesRepaintAndForcesFullPresent()
{
    {
        using namespace RedSalamander::DxUi;

        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* label = root->AddChild<Label>(L"Device recovery");
        label->SetBounds(D2D1::RectF(8.0f, 8.0f, 220.0f, 36.0f));
        window.Host().SetRoot(std::move(root));

        const WindowHostBitmapCapture initialCapture =
            CaptureAttachedHostWindowBitmapForWindowHostSuite(window, "device-loss initial render creates resources");
        Require(initialCapture.widthPx > 0u && initialCapture.heightPx > 0u, "device-loss initial render produces pixels");
        Require(window.Host().DebugHasD2DContext(), "device-loss initial render creates a D2D context");

        const uint64_t invalidatesBeforeDeviceLoss = window.Host().DebugGetInvalidateCount();
        window.Host().DebugSimulateDeviceLoss();
        Require(! window.Host().DebugHasD2DContext(), "device-loss simulation discards the D2D context");
        Require(window.Host().DebugGetInvalidateCount() > invalidatesBeforeDeviceLoss, "device-loss simulation schedules a repaint");

        const WindowHostBitmapCapture recoveredCapture =
            CaptureAttachedHostWindowBitmapForWindowHostSuite(window, "device-loss recovery render recreates resources");
        Require(recoveredCapture.widthPx > 0u && recoveredCapture.heightPx > 0u, "device-loss recovery render produces pixels");
        Require(window.Host().DebugHasD2DContext(), "device-loss recovery render recreates the D2D context");
    }

    const std::filesystem::path headerPath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.h";
    std::ifstream headerInput(headerPath);
    Require(headerInput.good(), "WindowHost header is readable for device-loss recovery guard");
    const std::string header((std::istreambuf_iterator<char>(headerInput)), std::istreambuf_iterator<char>());
    Require(header.find("_forceFullPresentAfterDeviceRecreate") != std::string::npos, "WindowHost stores a first-frame-after-recreate full-present flag");

    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.WindowHost.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "WindowHost source is readable for device-loss recovery guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t firstRender = source.find("void WindowHost::Render(const RECT* dirtyRectPx, bool allowHidden)");
    const size_t testRender  = source.find("void WindowHost::Render(const RECT* dirtyRectPx, WindowHostBitmapCapture* capture", firstRender);
    Require(firstRender != std::string::npos && testRender != std::string::npos && firstRender < testRender, "retail WindowHost Render overload is found");
    const std::string retailRenderBlock = source.substr(firstRender, testRender - firstRender);

    const size_t testRenderEnd = source.find("void WindowHost::OnSize", testRender);
    Require(testRenderEnd != std::string::npos && testRender < testRenderEnd, "test WindowHost Render overload is found");
    const std::string testRenderBlock = source.substr(testRender, testRenderEnd - testRender);

    for (const std::string& renderBlock : {retailRenderBlock, testRenderBlock})
    {
        const size_t partialDirty = renderBlock.find("dirtyRectMetrics.isPartialDirty && ! forceFullPresentAfterDeviceRecreate");
        Require(partialDirty != std::string::npos, "WindowHost Render suppresses partial present for the first recreated frame");

        const size_t deviceLost = renderBlock.find("render-device-lost");
        Require(deviceLost != std::string::npos, "WindowHost Render device-loss branch is found");
        const size_t resetShared = renderBlock.find("ResetSharedWindowHostGraphicsResources();", deviceLost);
        const size_t forceFull   = renderBlock.find("_forceFullPresentAfterDeviceRecreate = true;", deviceLost);
        const size_t invalidate  = renderBlock.find("Invalidate();", deviceLost);
        const size_t returnOut   = renderBlock.find("return;", deviceLost);
        Require(resetShared != std::string::npos && forceFull != std::string::npos && invalidate != std::string::npos && returnOut != std::string::npos,
                "WindowHost Render device-loss branch resets resources, forces full present, invalidates, and returns");
        Require(resetShared < forceFull && forceFull < invalidate && invalidate < returnOut,
                "WindowHost Render schedules full repaint after resource discard before returning from device loss");
    }
}

void TestWindowHostRenderSurvivesForcedNullSolidBrushes()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root = std::make_unique<Panel>();

    auto* menuBar = root->AddChild<MenuBar>();
    menuBar->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 32.0f));
    menuBar->SetItems({MenuBarItem{.text = L"File", .mnemonic = L'F'}, MenuBarItem{.text = L"View", .mnemonic = L'V'}});

    auto* toolbar = root->AddChild<Toolbar>();
    toolbar->SetBounds(D2D1::RectF(0.0f, 32.0f, 320.0f, 64.0f));
    toolbar->AddButton(L"Copy", L"\xE8C8");
    toolbar->AddSeparator();
    toolbar->AddToggleButton(L"Bold", L"\xE8DD");

    auto* button = root->AddChild<Button>(L"Apply");
    button->SetBounds(D2D1::RectF(12.0f, 76.0f, 120.0f, 108.0f));

    auto* toggle = root->AddChild<Toggle>(L"Feature");
    toggle->SetBounds(D2D1::RectF(132.0f, 76.0f, 308.0f, 108.0f));
    toggle->SetChecked(true);

    auto* checkbox = root->AddChild<Checkbox>(L"Remember");
    checkbox->SetBounds(D2D1::RectF(12.0f, 116.0f, 160.0f, 144.0f));
    checkbox->SetChecked(true);

    auto* radio = root->AddChild<RadioButton>(L"Daily");
    radio->SetBounds(D2D1::RectF(176.0f, 116.0f, 308.0f, 144.0f));
    radio->SetChecked(true);

    auto* slider = root->AddChild<Slider>();
    slider->SetBounds(D2D1::RectF(12.0f, 156.0f, 220.0f, 184.0f));
    slider->SetTickMarks({0.0, 50.0, 100.0});
    slider->SetValue(50.0);

    auto* status = root->AddChild<StatusStrip>();
    status->SetBounds(D2D1::RectF(0.0f, 186.0f, 320.0f, 208.0f));
    status->SetSections({StatusStrip::Section{.text = L"Ready", .widthDip = 0.0f}, StatusStrip::Section{.text = L"UTF-8", .widthDip = 80.0f}});

    auto* tabs = root->AddChild<TabControl>();
    tabs->AddTab<Label>(L"General", L"Pane");
    tabs->SetBounds(D2D1::RectF(0.0f, 208.0f, 320.0f, 320.0f));

    window.Host().SetSmokeOverlayVisible(true);
    window.Host().SetRoot(std::move(root));
    window.Host().DebugSetForceNullSolidBrushes(true);

    const uint64_t presentFailuresBefore = window.Host().DebugGetPresentFailureCount();
    const RedSalamander::DxUi::WindowHostBitmapCapture capture =
        CaptureAttachedHostWindowBitmapForWindowHostSuite(window, "forced-null solid brush render completes without crashing");
    Require(capture.widthPx > 0u && capture.heightPx > 0u, "forced-null solid brush render still produces a capture");
    Require(window.Host().DebugGetPresentFailureCount() == presentFailuresBefore, "forced-null solid brush render does not introduce present failures");

    window.Host().DebugSetForceNullSolidBrushes(false);
}

void TestWindowHostSmokeOverlayRendersBelowRootOverlay()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetSmokeOverlayVisible(true);
    window.Host().SetRoot(std::make_unique<OverlayZOrderControl>());

    const WindowHostBitmapCapture capture = CaptureAttachedHostWindowBitmapForWindowHostSuite(window, "smoke-overlay z-order capture succeeds");
    Require(capture.widthPx > 0u && capture.heightPx > 0u && ! capture.bgraPixels.empty(), "smoke-overlay z-order capture has pixels");

    const float dpi        = window.Host().GetDpi();
    const UINT overlayX    = DipToPixelForWindowHost(24.0f, dpi, capture.widthPx);
    const UINT overlayY    = DipToPixelForWindowHost(24.0f, dpi, capture.heightPx);
    const UINT contentX    = DipToPixelForWindowHost(96.0f, dpi, capture.widthPx);
    const UINT contentY    = DipToPixelForWindowHost(24.0f, dpi, capture.heightPx);
    const uint32_t overlay = GetWindowHostCapturePixelBgra(capture, overlayX, overlayY);
    const uint32_t content = GetWindowHostCapturePixelBgra(capture, contentX, contentY);

    const auto red   = [](uint32_t bgra) noexcept { return static_cast<uint8_t>((bgra >> 16u) & 0xFFu); };
    const auto green = [](uint32_t bgra) noexcept { return static_cast<uint8_t>((bgra >> 8u) & 0xFFu); };
    const auto blue  = [](uint32_t bgra) noexcept { return static_cast<uint8_t>(bgra & 0xFFu); };

    Require(green(overlay) > 180u && red(overlay) < 96u && blue(overlay) < 96u, "root overlay paints above the smoke overlay instead of being dimmed by it");
    Require(red(content) > (green(content) + 40u) && red(content) < 190u, "smoke overlay dims normal root content below root overlay paint");
}

void TestWindowHostOverlayHitTestingPrecedesContentHitTesting()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TrackingControlState overlayState;
    TrackingControlState contentState;

    auto root     = std::make_unique<Panel>();
    auto* overlay = root->AddChild<OverlayHitRecordingControl>(overlayState);
    auto* content = root->AddChild<TrackingControl>(contentState);

    root->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 80.0f));
    overlay->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 40.0f));
    content->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 40.0f));
    host.SetRoot(std::move(root));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(16, 16), handled));
    Require(handled, "overlay hit-test mouse-down is handled");
    Require(overlayState.mouseDownCount == 1u, "overlay hit target receives mouse-down before overlapping content");
    Require(contentState.mouseDownCount == 0u, "overlapping content is not invoked when an overlay hit target exists");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(16, 16), handled));
}

void TestWindowHostEscapeClosesMouseOpenedComboPopupBeforeCancelButton()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root          = std::make_unique<Panel>();
    auto* combo        = root->AddChild<ComboBox>();
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    cancelButton->SetBounds(D2D1::RectF(0.0f, 36.0f, 100.0f, 64.0f));
    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}});

    size_t cancelCount = 0u;
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    host.SetRoot(std::move(root));
    host.SetCancelButton(cancelButton);
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 120.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(172, 12), handled));
    Require(handled, "mouse-open combo click handled");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(172, 12), handled));
    Require(combo->GetHitBounds().bottom > combo->GetBounds().bottom, "mouse-opened combo popup is open before escape");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_ESCAPE, 0, handled));
    Require(handled, "escape handled while mouse-opened combo popup is open");
    Require(cancelCount == 0u, "cancel button not invoked while mouse-opened popup owns escape");
    Require(combo->GetHitBounds().bottom == combo->GetBounds().bottom, "escape closes mouse-opened combo popup first");
}

void TestWindowHostTabTraversalIncludesComboBox()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"First");
    auto* combo  = root->AddChild<ComboBox>();
    auto* field  = root->AddChild<TextField>();
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 80.0f, 24.0f));
    combo->SetBounds(D2D1::RectF(0.0f, 28.0f, 120.0f, 56.0f));
    field->SetBounds(D2D1::RectF(0.0f, 60.0f, 140.0f, 88.0f));
    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}});

    host.SetRoot(std::move(root));
    host.SetFocusControl(button);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab to combo handled");
    Require(host.GetFocusControl() == combo, "combo participates in tab traversal");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from combo handled");
    Require(host.GetFocusControl() == field, "tab advances from combo to next focusable control");
}

void TestWindowHostTabTraversalStaysConsistentAcrossFieldComboTreeGridAndButtons()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root      = std::make_unique<Panel>();
    auto* field    = root->AddChild<TextField>();
    auto* combo    = root->AddChild<ComboBox>();
    auto* tree     = root->AddChild<Tree>();
    auto* grid     = root->AddChild<Grid>();
    auto* okButton = root->AddChild<Button>(L"OK");

    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 28.0f));
    combo->SetBounds(D2D1::RectF(0.0f, 36.0f, 160.0f, 64.0f));
    tree->SetBounds(D2D1::RectF(0.0f, 72.0f, 220.0f, 156.0f));
    grid->SetBounds(D2D1::RectF(0.0f, 164.0f, 260.0f, 252.0f));
    okButton->SetBounds(D2D1::RectF(0.0f, 260.0f, 96.0f, 288.0f));

    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}});

    MutableTreeModel treeModel;
    treeModel.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"General"},
        TreeItemData{.id = 2u, .text = L"Viewers"},
    });
    tree->SetModel(&treeModel);

    MultiRowGridModel gridModel(3u);
    RecordingGridDelegate gridDelegate;
    grid->SetModel(&gridModel);
    grid->SetDelegate(&gridDelegate);

    host.SetRoot(std::move(root));
    host.SetFocusControl(field);
    Require(host.GetFocusControl() == field, "mixed retained-tree traversal starts on the text field");

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from text field handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == combo, "tab advances from text field to combo");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from combo handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == tree, "tab advances from combo to tree");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from tree handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == grid, "tab advances from tree to grid");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from grid handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == okButton, "tab advances from grid to command button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from command button handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == field, "tab wraps from command button back to the text field");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "shift+tab from text field handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == okButton, "shift+tab wraps from text field to command button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "shift+tab from command button handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == grid, "shift+tab moves from command button to grid");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "shift+tab from grid handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == tree, "shift+tab moves from grid to tree");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "shift+tab from tree handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == combo, "shift+tab moves from tree to combo");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "shift+tab from combo handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == field, "shift+tab moves from combo back to text field");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYUP, VK_SHIFT, 0, handled));
}

void TestWindowHostGroupedListNavigationKeepsTreeTypeaheadAndGridSelectionVisible()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root         = std::make_unique<Panel>();
    auto* tree        = root->AddChild<Tree>();
    auto* grid        = root->AddChild<Grid>();
    auto* closeButton = root->AddChild<Button>(L"Close");

    tree->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 108.0f));
    grid->SetBounds(D2D1::RectF(0.0f, 116.0f, 280.0f, 292.0f));
    grid->SetRowHeightDip(24.0f);
    grid->SetHeaderHeightDip(32.0f);
    closeButton->SetBounds(D2D1::RectF(0.0f, 300.0f, 96.0f, 328.0f));

    MutableTreeModel treeModel;
    const auto setTreeItems = [&treeModel](bool expanded)
    {
        if (expanded)
        {
            treeModel.SetVisibleItems({
                TreeItemData{.id = 1u, .text = L"Plugins", .hasChildren = true, .expanded = true},
                TreeItemData{.id = 2u, .parentId = 1u, .text = L"ViewerSqlite", .depth = 1u},
                TreeItemData{.id = 3u, .text = L"Search"},
                TreeItemData{.id = 4u, .text = L"Themes"},
            });
            return;
        }

        treeModel.SetVisibleItems({
            TreeItemData{.id = 1u, .text = L"Plugins", .hasChildren = true, .expanded = false},
            TreeItemData{.id = 3u, .text = L"Search"},
            TreeItemData{.id = 4u, .text = L"Themes"},
        });
    };
    setTreeItems(false);

    RecordingTreeDelegate treeDelegate;
    tree->SetModel(&treeModel);
    tree->SetDelegate(&treeDelegate);

    GroupedGridModel gridModel(6u);
    gridModel.SetGroups({
        GroupedGridModel::Group{.stableId = 10u, .title = L"Favorites", .startRowIndex = 0u, .rowCount = 2u},
        GroupedGridModel::Group{.stableId = 20u, .title = L"Folders", .startRowIndex = 3u, .rowCount = 2u},
    });

    CollapsibleGroupedGridDelegate gridDelegate(gridModel);
    grid->SetModel(&gridModel);
    grid->SetDelegate(&gridDelegate);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 360.0f));

    tree->SetSelectedItemId(1u);
    grid->GetSelectionModel().SetSingle(gridModel.GetStableRowId(1u));
    host.SetFocusControl(tree);
    Require(host.GetFocusControl() == tree, "grouped/list host navigation starts with tree focus");

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RIGHT, 0, handled));
    Require(handled, "grouped/list host handles tree right key on a collapsed parent");
    Require(treeDelegate.toggleCount == 1u, "grouped/list host requests one tree expansion");
    Require(treeDelegate.lastToggledItemId == 1u && treeDelegate.lastExpandedState, "grouped/list host tree right key expands the parent item");

    setTreeItems(true);
    tree->NotifyDataChanged();

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RIGHT, 0, handled));
    Require(handled, "grouped/list host handles tree right key on an expanded parent");
    Require(tree->GetSelectedItemId().has_value() && tree->GetSelectedItemId().value() == 2u,
            "grouped/list host tree right key selects the first child when the parent is expanded");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_LEFT, 0, handled));
    Require(handled, "grouped/list host handles tree left key on a child item");
    Require(tree->GetSelectedItemId().has_value() && tree->GetSelectedItemId().value() == 1u,
            "grouped/list host tree left key selects the parent from the child row");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_LEFT, 0, handled));
    Require(handled, "grouped/list host handles tree left key on an expanded parent");
    Require(treeDelegate.toggleCount == 2u, "grouped/list host requests one tree collapse after expansion");
    Require(treeDelegate.lastToggledItemId == 1u && ! treeDelegate.lastExpandedState, "grouped/list host tree left key collapses the parent item");

    setTreeItems(false);
    tree->NotifyDataChanged();

    const size_t selectionChangedBeforeTypeahead = treeDelegate.selectionChangedCount;
    handled                                      = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_CHAR, static_cast<WPARAM>(L's'), 0, handled));
    Require(handled, "grouped/list host handles tree typeahead");
    Require(tree->GetSelectedItemId().has_value() && tree->GetSelectedItemId().value() == 3u,
            "grouped/list host tree typeahead selects the matching visible item after collapse");
    Require(treeDelegate.selectionChangedCount == selectionChangedBeforeTypeahead + 1u,
            "grouped/list host tree typeahead notifies one visible selection change");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "grouped/list host handles tab from tree to grouped grid");
    Require(host.GetFocusControl() == grid, "grouped/list host tab advances focus from tree to grouped grid");

    Require(grid->GetSelectionModel().GetCount() == 1u, "grouped/list host grouped grid starts with one selected row");
    Require(grid->GetSelectionModel().GetOrderedSelection().front() == gridModel.GetStableRowId(1u),
            "grouped/list host grouped grid starts on a row that will become hidden when the group collapses");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_LEFT, 0, handled));
    Require(handled, "grouped/list host grouped grid handles left key to collapse the selected group");
    Require(gridDelegate.groupToggleCount == 1u, "grouped/list host grouped grid reports one collapse toggle");
    Require(gridDelegate.lastGroupStableId == 10u && gridDelegate.lastGroupCollapsed,
            "grouped/list host grouped grid reports the collapsed group id and state from the keyboard path");
    Require(grid->GetSelectionModel().GetCount() == 1u, "grouped/list host grouped grid keeps one visible row selected after collapse");
    Require(grid->GetSelectionModel().GetOrderedSelection().front() == gridModel.GetStableRowId(2u),
            "grouped/list host grouped grid rehomes selection to the nearest visible row after collapse");

    const GridVisibleWorkMetrics collapsedMetrics = grid->GetVisibleWorkMetrics();
    Require(collapsedMetrics.visibleRowCount == 4u, "grouped/list host grouped grid collapse updates visible row work");
    Require(collapsedMetrics.visibleGroupHeaderCount == 2u, "grouped/list host grouped grid keeps visible headers after collapse");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RIGHT, 0, handled));
    Require(handled, "grouped/list host grouped grid handles right key to re-expand the collapsed group from the fallback row");
    Require(gridDelegate.groupToggleCount == 2u, "grouped/list host grouped grid reports one expand toggle after keyboard collapse");
    Require(gridDelegate.lastGroupStableId == 10u && ! gridDelegate.lastGroupCollapsed,
            "grouped/list host grouped grid reports the expanded group id and state from the keyboard path");
    Require(grid->GetSelectionModel().GetCount() == 1u, "grouped/list host grouped grid keeps one selected row after re-expansion");
    Require(grid->GetSelectionModel().GetOrderedSelection().front() == gridModel.GetStableRowId(2u),
            "grouped/list host grouped grid keeps the fallback row selected after re-expansion");

    const GridVisibleWorkMetrics expandedMetrics = grid->GetVisibleWorkMetrics();
    Require(expandedMetrics.visibleRowCount == 4u, "grouped/list host grouped grid re-expansion restores visible row work");
    Require(expandedMetrics.visibleGroupHeaderCount == 2u, "grouped/list host grouped grid keeps visible headers after re-expansion");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_END, 0, handled));
    Require(handled, "grouped/list host handles grouped grid end key after keyboard collapse and re-expansion");
    Require(grid->GetSelectionModel().GetOrderedSelection().front() == gridModel.GetStableRowId(5u),
            "grouped/list host grouped grid end key keeps keyboard navigation on the last visible row after re-expansion");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "grouped/list host handles tab from grouped grid to the dialog button");
    Require(host.GetFocusControl() == closeButton, "grouped/list host tab advances from the grouped grid to the dialog button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));
    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "grouped/list host handles shift+tab from the dialog button back to the grouped grid");
    Require(host.GetFocusControl() == grid, "grouped/list host shift+tab returns focus to the grouped grid");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYUP, VK_SHIFT, 0, handled));
}

void TestWindowHostAltDownOpensComboPopup()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}});

    host.SetRoot(std::move(root));
    host.SetFocusControl(combo);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSKEYDOWN, VK_DOWN, 0, handled));
    Require(handled, "alt+down is handled for combo");
    Require(combo->GetHitBounds().bottom > combo->GetBounds().bottom, "alt+down opens combo popup");
}

void TestWindowHostAltUpClosesComboPopup()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}});

    host.SetRoot(std::move(root));
    host.SetFocusControl(combo);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSKEYDOWN, VK_DOWN, 0, handled));
    Require(handled, "alt+down handled before alt+up close");
    Require(combo->GetHitBounds().bottom > combo->GetBounds().bottom, "combo popup is open before alt+up");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSKEYDOWN, VK_UP, 0, handled));
    Require(handled, "alt+up is handled for open combo");
    Require(combo->GetHitBounds().bottom == combo->GetBounds().bottom, "alt+up closes combo popup");
}

void TestWindowHostMnemonicActivatesButton()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Find");
    button->SetMnemonic(L'F');
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));

    bool clicked = false;
    button->SetOnClick([&clicked] { clicked = true; });
    host.SetRoot(std::move(root));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSCHAR, L'f', 0, handled));
    Require(handled, "button mnemonic handled");
    Require(clicked, "button mnemonic invokes click");
    Require(host.GetFocusControl() == button, "button mnemonic focuses button");
}

void TestWindowHostLabelMnemonicTargetsField()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* label = root->AddChild<Label>(L"Root");
    auto* combo = root->AddChild<ComboBox>();
    label->SetMnemonic(L'R');
    label->SetMnemonicTarget(combo);
    combo->SetEditable(true);
    label->SetBounds(D2D1::RectF(0.0f, 0.0f, 80.0f, 24.0f));
    combo->SetBounds(D2D1::RectF(0.0f, 28.0f, 180.0f, 56.0f));

    host.SetRoot(std::move(root));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSCHAR, L'r', 0, handled));
    Require(handled, "label mnemonic handled");
    Require(host.GetFocusControl() == combo, "label mnemonic focuses target control");
}

void TestWindowHostUnknownMnemonicRemainsUnhandled()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Open");
    button->SetMnemonic(L'O');
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    host.SetRoot(std::move(root));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSCHAR, L'x', 0, handled));
    Require(! handled, "unknown mnemonic remains unhandled");
    Require(host.GetFocusControl() == nullptr, "unknown mnemonic does not move focus");
}

void TestWindowHostHiddenAnimationTickDropsSubscriptionUntilShown()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root     = std::make_unique<Panel>();
    auto* ticking = root->AddChild<AnimationTickTraceControl>();
    window.Host().SetRoot(std::move(root));
    window.Host().RequestAnimation();

    Require(IsWindowVisible(window.Hwnd()) == FALSE, "attached host window starts hidden for hidden-animation test");
    Require(window.Host().DebugHasActiveAnimationSubscription(), "hidden-animation test starts with an active subscription");
    const uint64_t invalidatesBefore = window.Host().DebugGetInvalidateCount();

    Require(! window.Host().DebugAnimationTickForTest(123u), "hidden host animation tick releases the dispatcher subscription");
    Require(! window.Host().DebugHasActiveAnimationSubscription(),
            "hidden host animation tick clears the subscription while remembering the suspended animation");
    Require(ticking->tickCount == 0u, "hidden host animation tick skips root ticking while hidden");
    Require(window.Host().DebugGetInvalidateCount() == invalidatesBefore, "hidden host animation tick skips invalidation while hidden");

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    Require(window.Host().DebugHasActiveAnimationSubscription(), "showing the host restores a suspended animation subscription");
    Require(window.Host().DebugAnimationTickForTest(140u), "visible host animation tick resumes root ticking after show");
    Require(ticking->tickCount == 1u, "visible host animation tick reaches the root after show");
}

void TestWindowHostRestoreFromMinimizeRearmsSuspendedAnimation()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root = std::make_unique<Panel>();
    static_cast<void>(root->AddChild<AnimationTickTraceControl>());
    window.Host().SetRoot(std::move(root));
    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.Host().RequestAnimation();

    ShowWindow(window.Hwnd(), SW_MINIMIZE);
    Require(! window.Host().DebugAnimationTickForTest(200u), "iconic host latches its animation as suspended");
    Require(! window.Host().DebugHasActiveAnimationSubscription(), "iconic host releases its animation subscription");

    ShowWindow(window.Hwnd(), SW_RESTORE);
    Require(window.Host().DebugHasActiveAnimationSubscription(), "SIZE_RESTORED re-arms a suspended animation for an effectively visible host");
}

} // namespace

void RunWindowHostTests()
{
    auto runTest = [](const char* name, void (*fn)())
    {
        std::cerr << "  [START] " << name << '\n' << std::flush;
        fn();
        std::cerr << "  [DONE] " << name << '\n' << std::flush;
    };

    runTest("TestDxUiTypographyMapsFontRolesToSegoeUiVariableFamilies", TestDxUiTypographyMapsFontRolesToSegoeUiVariableFamilies);
    runTest("TestDxUiTypographyMeasurementCachesFormatsAndFamilyResolution", TestDxUiTypographyMeasurementCachesFormatsAndFamilyResolution);
    runTest("TestConnectionCredentialPromptDestroysWindowOnModalQuit", TestConnectionCredentialPromptDestroysWindowOnModalQuit);
    runTest("TestConnectionCredentialPromptTeardownDoesNotWipeThroughRawTextFieldPointers",
            TestConnectionCredentialPromptTeardownDoesNotWipeThroughRawTextFieldPointers);
    runTest("TestConnectionCredentialPromptUiaPumpDoesNotDetachTimedOutWorkers", TestConnectionCredentialPromptUiaPumpDoesNotDetachTimedOutWorkers);
    runTest("TestWindowHostPaintUsesWilBeginPaintRaii", TestWindowHostPaintUsesWilBeginPaintRaii);
    runTest("TestWindowHostKeyboardInputMarksFocusVisible", TestWindowHostKeyboardInputMarksFocusVisible);
    runTest("TestWindowHostPointerInputClearsKeyboardFocusVisible", TestWindowHostPointerInputClearsKeyboardFocusVisible);
    runTest("TestWindowHostCrossThreadDetachReleasesOwnerThreadAttachmentCount", TestWindowHostCrossThreadDetachReleasesOwnerThreadAttachmentCount);
    runTest("TestWindowHostDetachKeepsSharedGraphicsAttachmentUntilControlTreeDestroyed",
            TestWindowHostDetachKeepsSharedGraphicsAttachmentUntilControlTreeDestroyed);
    runTest("TestWindowHostDestructorDetachesBeforeMemberTeardown", TestWindowHostDestructorDetachesBeforeMemberTeardown);
    runTest("TestWindowHostEmitsFrameStageMetricsForCaptureRender", TestWindowHostEmitsFrameStageMetricsForCaptureRender);
    runTest("TestWindowHostBlocksLayoutMutationDuringRender", TestWindowHostBlocksLayoutMutationDuringRender);
    runTest("TestPostMessagePayloadTeardownDrainDeletesUndeliveredPayloads", TestPostMessagePayloadTeardownDrainDeletesUndeliveredPayloads);
    runTest("TestContiguousPostedPayloadCoalescingPreservesQueueOrderAndOperationKeys",
            TestContiguousPostedPayloadCoalescingPreservesQueueOrderAndOperationKeys);
    runTest("TestSharedTestSupportPreservesSandboxAndEnvironmentPolicies", TestSharedTestSupportPreservesSandboxAndEnvironmentPolicies);
    runTest("TestSharedTestSupportPumpsMessagesAndBoundsSnapshotPolling", TestSharedTestSupportPumpsMessagesAndBoundsSnapshotPolling);
    runTest("TestWindowHostMouseMoveUpdatesHoverTarget", TestWindowHostMouseMoveUpdatesHoverTarget);
    runTest("TestWindowHostMouseLeaveOverForeignPopupClearsHover", TestWindowHostMouseLeaveOverForeignPopupClearsHover);
    runTest("TestWindowHostMouseLeaveWithForeignCaptureClearsHover", TestWindowHostMouseLeaveWithForeignCaptureClearsHover);
    runTest("TestWindowHostTabTraversal", TestWindowHostTabTraversal);
    runTest("TestWindowHostShiftTabTraversal", TestWindowHostShiftTabTraversal);
    runTest("TestWindowHostNativeFocusLossRetainsLogicalFocusForTraversal", TestWindowHostNativeFocusLossRetainsLogicalFocusForTraversal);
    runTest("TestWindowHostReturnInvokesDefaultButtonWhenFocusedControlDoesNotOwnEnter",
            TestWindowHostReturnInvokesDefaultButtonWhenFocusedControlDoesNotOwnEnter);
    runTest("TestWindowHostReturnInvokesDefaultButtonWhenNoControlIsFocused", TestWindowHostReturnInvokesDefaultButtonWhenNoControlIsFocused);
    runTest("TestWindowHostReturnDoesNotInvokeDefaultButtonWhenFocusedControlOwnsEnter",
            TestWindowHostReturnDoesNotInvokeDefaultButtonWhenFocusedControlOwnsEnter);
    runTest("TestButtonKeyboardActivationCanReplaceRootSafely", TestButtonKeyboardActivationCanReplaceRootSafely);
    runTest("TestWindowHostSpaceAndReturnInvokeFocusedButtonWithoutDefaultButtonFallback",
            TestWindowHostSpaceAndReturnInvokeFocusedButtonWithoutDefaultButtonFallback);
    runTest("TestWindowHostDpiChangedIsHandled", TestWindowHostDpiChangedIsHandled);
    runTest("TestWindowHostDpiChangedInvalidatesMultilineCachesAndResizesAttachedWindow",
            TestWindowHostDpiChangedInvalidatesMultilineCachesAndResizesAttachedWindow);
    runTest("TestWindowHostDpiChangedRecreatesTargetBitmapWhenSizeIsUnchanged", TestWindowHostDpiChangedRecreatesTargetBitmapWhenSizeIsUnchanged);
    runTest("TestWindowHostAttachedWindowsRenderAcrossUiThreads", TestWindowHostAttachedWindowsRenderAcrossUiThreads);
    runTest("TestWindowHostEscapeInvokesCancelButton", TestWindowHostEscapeInvokesCancelButton);
    runTest("TestWindowHostEscapeClosesComboPopupBeforeCancelButton", TestWindowHostEscapeClosesComboPopupBeforeCancelButton);
    runTest("TestWindowHostMenuKeyInvokesFocusedButtonContextMenu", TestWindowHostMenuKeyInvokesFocusedButtonContextMenu);
    runTest("TestWindowHostShiftF10InvokesFocusedToggleContextMenu", TestWindowHostShiftF10InvokesFocusedToggleContextMenu);
    runTest("TestWindowHostMenuKeyInvokesFocusedCheckboxContextMenu", TestWindowHostMenuKeyInvokesFocusedCheckboxContextMenu);
    runTest("TestWindowHostSpaceAndReturnToggleFocusedToggleWithoutDefaultButtonFallback",
            TestWindowHostSpaceAndReturnToggleFocusedToggleWithoutDefaultButtonFallback);
    runTest("TestWindowHostSpaceTogglesFocusedCheckboxAndReturnInvokesDefaultButton", TestWindowHostSpaceTogglesFocusedCheckboxAndReturnInvokesDefaultButton);
    runTest("TestWindowHostMixedDialogKeyboardFlowKeepsCommandsOnFocusedControls", TestWindowHostMixedDialogKeyboardFlowKeepsCommandsOnFocusedControls);
    runTest("TestWindowHostMixedDialogMouseFlowKeepsCommandsOnHitControls", TestWindowHostMixedDialogMouseFlowKeepsCommandsOnHitControls);
    runTest("TestWindowHostMenuKeyInvokesFocusedTreeContextMenu", TestWindowHostMenuKeyInvokesFocusedTreeContextMenu);
    runTest("TestWindowHostShiftF10InvokesFocusedTreeContextMenu", TestWindowHostShiftF10InvokesFocusedTreeContextMenu);
    runTest("TestWindowHostMenuKeyInvokesFocusedGridContextMenu", TestWindowHostMenuKeyInvokesFocusedGridContextMenu);
    runTest("TestWindowHostShiftF10InvokesFocusedGridContextMenu", TestWindowHostShiftF10InvokesFocusedGridContextMenu);
    runTest("TestWindowHostMenuKeyInvokesFocusedTextFieldContextMenu", TestWindowHostMenuKeyInvokesFocusedTextFieldContextMenu);
    runTest("TestWindowHostMenuKeyInvokesFocusedComboContextMenu", TestWindowHostMenuKeyInvokesFocusedComboContextMenu);
    runTest("TestWindowHostSetRootClearsDestroyedTreeInteractionState", TestWindowHostSetRootClearsDestroyedTreeInteractionState);
    runTest("TestWindowHostDetachDeactivatesSecureTextInputBeforeDestroyingRoot", TestWindowHostDetachDeactivatesSecureTextInputBeforeDestroyingRoot);
    runTest("TestWindowHostClearChildrenPrunesDestroyedTreeInteractionState", TestWindowHostClearChildrenPrunesDestroyedTreeInteractionState);
    runTest("TestWindowHostKeyDownCallbackDisablingFocusPrunesBeforePostDispatchSync", TestWindowHostKeyDownCallbackDisablingFocusPrunesBeforePostDispatchSync);
    runTest("TestWindowHostCharCallbackDisablingFocusPrunesBeforePostDispatchSync", TestWindowHostCharCallbackDisablingFocusPrunesBeforePostDispatchSync);
    runTest("TestWindowHostFocusLossCallbackRevalidatesRequestedFocusTarget", TestWindowHostFocusLossCallbackRevalidatesRequestedFocusTarget);
    runTest("TestWindowHostIgnoresObserverButtonsOutsideInstalledRoot", TestWindowHostIgnoresObserverButtonsOutsideInstalledRoot);
    runTest("TestWindowHostIgnoresFocusAndCaptureOutsideInstalledRoot", TestWindowHostIgnoresFocusAndCaptureOutsideInstalledRoot);
    runTest("TestWindowHostCaptureLossClearsPressedButtonState", TestWindowHostCaptureLossClearsPressedButtonState);
    runTest("TestWindowHostRedundantCaptureDoesNotCancelMouseDownCapture", TestWindowHostRedundantCaptureDoesNotCancelMouseDownCapture);
    runTest("TestWindowHostPointerDispatchDoesNotReuseTargetAfterRootReplacement", TestWindowHostPointerDispatchDoesNotReuseTargetAfterRootReplacement);
    runTest("TestWindowHostHoverEnterDoesNotReuseTargetAfterRootReplacement", TestWindowHostHoverEnterDoesNotReuseTargetAfterRootReplacement);
    runTest("TestWindowHostDeviceLossSchedulesRepaintAndForcesFullPresent", TestWindowHostDeviceLossSchedulesRepaintAndForcesFullPresent);
    runTest("TestWindowHostRenderSurvivesForcedNullSolidBrushes", TestWindowHostRenderSurvivesForcedNullSolidBrushes);
    runTest("TestWindowHostSmokeOverlayRendersBelowRootOverlay", TestWindowHostSmokeOverlayRendersBelowRootOverlay);
    runTest("TestWindowHostOverlayHitTestingPrecedesContentHitTesting", TestWindowHostOverlayHitTestingPrecedesContentHitTesting);
    runTest("TestWindowHostEscapeClosesMouseOpenedComboPopupBeforeCancelButton", TestWindowHostEscapeClosesMouseOpenedComboPopupBeforeCancelButton);
    runTest("TestWindowHostTabTraversalIncludesComboBox", TestWindowHostTabTraversalIncludesComboBox);
    runTest("TestWindowHostTabTraversalStaysConsistentAcrossFieldComboTreeGridAndButtons",
            TestWindowHostTabTraversalStaysConsistentAcrossFieldComboTreeGridAndButtons);
    runTest("TestWindowHostGroupedListNavigationKeepsTreeTypeaheadAndGridSelectionVisible",
            TestWindowHostGroupedListNavigationKeepsTreeTypeaheadAndGridSelectionVisible);
    runTest("TestWindowHostAltDownOpensComboPopup", TestWindowHostAltDownOpensComboPopup);
    runTest("TestWindowHostAltUpClosesComboPopup", TestWindowHostAltUpClosesComboPopup);
    runTest("TestWindowHostMnemonicActivatesButton", TestWindowHostMnemonicActivatesButton);
    runTest("TestWindowHostLabelMnemonicTargetsField", TestWindowHostLabelMnemonicTargetsField);
    runTest("TestWindowHostUnknownMnemonicRemainsUnhandled", TestWindowHostUnknownMnemonicRemainsUnhandled);
    runTest("TestWindowHostHiddenAnimationTickDropsSubscriptionUntilShown", TestWindowHostHiddenAnimationTickDropsSubscriptionUntilShown);
    runTest("TestWindowHostRestoreFromMinimizeRearmsSuspendedAnimation", TestWindowHostRestoreFromMinimizeRearmsSuspendedAnimation);
}

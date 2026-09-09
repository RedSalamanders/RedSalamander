#include "Terminal.h"

#if defined(ENABLE_TESTS)

#include <algorithm>
#include <array>
#include <chrono>
#include <format>
#include <iterator>
#include <limits>
#include <span>
#include <string>
#include <string_view>
#include <thread>
#include <vector>

#include <bcrypt.h>
#include <psapi.h>

#pragma comment(lib, "bcrypt")
#pragma comment(lib, "psapi")

#include "Helpers.h"
#include "TerminalInternal.h"

extern HINSTANCE g_hInstance;

using namespace TerminalPluginDetail;

namespace
{

struct VtUpgradeCorpus final
{
    std::string prelude;
    std::string kittyPartial;
    std::string kittyCompletion;
    std::string postlude;
    std::string postludeWithoutExpectedDelta;
    std::string canonicalActionStream;
    uint64_t inputByteCount = 0u;
};

[[nodiscard]] std::string EncodeRepeatedByteBase64(size_t byteCount, uint8_t value)
{
    constexpr std::string_view alphabet =
        "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    std::string encoded;
    encoded.reserve(((byteCount + 2u) / 3u) * 4u);
    size_t remaining = byteCount;
    while (remaining >= 3u)
    {
        const uint32_t packed = (static_cast<uint32_t>(value) << 16u) |
            (static_cast<uint32_t>(value) << 8u) | static_cast<uint32_t>(value);
        encoded.push_back(alphabet[(packed >> 18u) & 0x3Fu]);
        encoded.push_back(alphabet[(packed >> 12u) & 0x3Fu]);
        encoded.push_back(alphabet[(packed >> 6u) & 0x3Fu]);
        encoded.push_back(alphabet[packed & 0x3Fu]);
        remaining -= 3u;
    }
    if (remaining != 0u)
    {
        const uint32_t packed = (static_cast<uint32_t>(value) << 16u) |
            (remaining == 2u ? static_cast<uint32_t>(value) << 8u : 0u);
        encoded.push_back(alphabet[(packed >> 18u) & 0x3Fu]);
        encoded.push_back(alphabet[(packed >> 12u) & 0x3Fu]);
        encoded.push_back(remaining == 2u ? alphabet[(packed >> 6u) & 0x3Fu] : '=');
        encoded.push_back('=');
    }
    return encoded;
}

[[nodiscard]] std::string MakeOsc52(size_t decodedBytes)
{
    std::string sequence = "\x1b]52;c;";
    sequence.append(EncodeRepeatedByteBase64(decodedBytes, 0x78u));
    sequence.push_back('\a');
    return sequence;
}

void AppendCorpusWrite(std::string& canonical, std::string_view bytes)
{
    canonical.append(std::format("write:{}\n", bytes.size()));
    canonical.append(bytes);
    canonical.push_back('\n');
}

void AppendCorpusAction(std::string& canonical, std::string_view action)
{
    canonical.append(std::format("action:{}:{}\n", action.size(), action));
}

[[nodiscard]] VtUpgradeCorpus BuildVtUpgradeCorpus()
{
    VtUpgradeCorpus corpus;
    corpus.prelude =
        "\x1b" "c"
        "\x1b_Ga=d,d=A,q=2;\x1b\\"
        "RedSalamander VT upgrade baseline: ASCII | UTF-8: \xE7\xAB\xAF\xE6\x9C\xAB "
        "\xF0\x9F\xA7\xAA | wide: \xE7\x95\x8C\r\n"
        "\x1b[1;31mred-bold\x1b[0m \x1b[38;2;12;34;56mtruecolor\x1b[0m "
        "\x1b]8;id=upgrade;https://example.invalid/libghostty\x1b\\link\x1b]8;;\x1b\\\r\n"
        "\x1b[4;9Hcursor\x1b[2K\x1b[2;12r\x1b[2S\x1b[1T\x1b[r"
        "\x1b[?1049halt-screen\x1b[3;4Halt\x1b[?1049l";
    corpus.prelude.append(MakeOsc52((256u * 1024u) - 1u));
    corpus.prelude.append(MakeOsc52(256u * 1024u));
    corpus.prelude.append(MakeOsc52((256u * 1024u) + 1u));
    corpus.kittyPartial = "\x1b_Gf=32,s=1,v=1,i=77,a=T,m=1;/wAA\x1b\\";
    corpus.kittyCompletion = "\x1b_Gi=77,m=0;/w==\x1b\\";
    corpus.postludeWithoutExpectedDelta = "post-kitty\r\n\x1b[5n\x1b[6n\x1b[c";
    corpus.postlude = "post-kitty\r\n\x1bPq";
    corpus.postlude.push_back(static_cast<char>(0x80u));
    corpus.postlude.push_back(static_cast<char>(0xFEu));
    corpus.postlude.push_back(static_cast<char>(0xFFu));
    corpus.postlude.append("\x1b\\\x1b[5n\x1b[6n\x1b[c");

    corpus.canonicalActionStream = "red-salamander-terminal-vt-upgrade-corpus-v1\n";
    AppendCorpusWrite(corpus.canonicalActionStream, corpus.prelude);
    AppendCorpusAction(corpus.canonicalActionStream, "resize:96x32:8x16");
    AppendCorpusWrite(corpus.canonicalActionStream, corpus.kittyPartial);
    AppendCorpusAction(corpus.canonicalActionStream, "snapshot:kitty-pending");
    AppendCorpusWrite(corpus.canonicalActionStream, corpus.kittyCompletion);
    AppendCorpusWrite(corpus.canonicalActionStream, corpus.postlude);
    AppendCorpusAction(corpus.canonicalActionStream, "resize:120x40:8x16");
    AppendCorpusAction(corpus.canonicalActionStream, "snapshot:semantic-v1:repeat-200");
    AppendCorpusAction(corpus.canonicalActionStream, "hosted-render:steady-state:repeat-200");
    corpus.inputByteCount = corpus.prelude.size() + corpus.kittyPartial.size() +
        corpus.kittyCompletion.size() + corpus.postlude.size();
    return corpus;
}

[[nodiscard]] bool ComputeSha256(
    std::span<const uint8_t> bytes,
    std::array<char, TerminalVtUpgradeTestContract::kSha256TextBytes>& digestText) noexcept
{
    // This is a test-only in-memory digest boundary. The product runtime file
    // identity path remains owned by TerminalRuntimeLoader.cpp.
    wil::unique_bcrypt_algorithm algorithm;
    NTSTATUS status = BCryptOpenAlgorithmProvider(algorithm.put(), BCRYPT_SHA256_ALGORITHM, nullptr, 0u);
    if (! BCRYPT_SUCCESS(status))
    {
        return false;
    }
    DWORD hashObjectBytes = 0u;
    DWORD propertyBytes = 0u;
    status = BCryptGetProperty(algorithm.get(),
                               BCRYPT_OBJECT_LENGTH,
                               reinterpret_cast<PUCHAR>(&hashObjectBytes),
                               sizeof(hashObjectBytes),
                               &propertyBytes,
                               0u);
    if (! BCRYPT_SUCCESS(status) || hashObjectBytes == 0u)
    {
        return false;
    }
    std::vector<std::byte> hashObject(hashObjectBytes);
    wil::unique_bcrypt_hash hash;
    status = BCryptCreateHash(
        algorithm.get(), hash.put(), reinterpret_cast<PUCHAR>(hashObject.data()), hashObjectBytes, nullptr, 0u, 0u);
    if (! BCRYPT_SUCCESS(status))
    {
        return false;
    }
    if (! bytes.empty())
    {
        if (bytes.size() > (std::numeric_limits<ULONG>::max)())
        {
            return false;
        }
        status = BCryptHashData(hash.get(),
                                reinterpret_cast<PUCHAR>(const_cast<uint8_t*>(bytes.data())),
                                static_cast<ULONG>(bytes.size()),
                                0u);
        if (! BCRYPT_SUCCESS(status))
        {
            return false;
        }
    }
    std::array<uint8_t, 32u> digest{};
    status = BCryptFinishHash(hash.get(), digest.data(), static_cast<ULONG>(digest.size()), 0u);
    if (! BCRYPT_SUCCESS(status))
    {
        return false;
    }
    constexpr char hex[] = "0123456789abcdef";
    for (size_t index = 0u; index < digest.size(); ++index)
    {
        digestText[index * 2u] = hex[digest[index] >> 4u];
        digestText[index * 2u + 1u] = hex[digest[index] & 0x0Fu];
    }
    digestText.back() = '\0';
    return true;
}

[[nodiscard]] bool ComputeSha256(
    std::string_view bytes,
    std::array<char, TerminalVtUpgradeTestContract::kSha256TextBytes>& digestText) noexcept
{
    return ComputeSha256(
        std::span(reinterpret_cast<const uint8_t*>(bytes.data()), bytes.size()), digestText);
}

void ObserveProcessMemory(TerminalVtUpgradeTestContract::Evidence& evidence) noexcept
{
    PROCESS_MEMORY_COUNTERS_EX counters{};
    counters.cb = sizeof(counters);
    if (GetProcessMemoryInfo(
            GetCurrentProcess(), reinterpret_cast<PROCESS_MEMORY_COUNTERS*>(&counters), sizeof(counters)) != FALSE)
    {
        evidence.peakPrivateBytes = std::max<uint64_t>(evidence.peakPrivateBytes, counters.PrivateUsage);
        evidence.peakWorkingSetBytes = std::max<uint64_t>(evidence.peakWorkingSetBytes, counters.PeakWorkingSetSize);
    }
}

[[nodiscard]] uint64_t ElapsedUs(std::chrono::steady_clock::time_point startedAt) noexcept
{
    const auto nanoseconds = std::chrono::duration_cast<std::chrono::nanoseconds>(
        std::chrono::steady_clock::now() - startedAt).count();
    return std::max<uint64_t>(1u, (static_cast<uint64_t>(std::max<int64_t>(0, nanoseconds)) + 999u) / 1'000u);
}

} // namespace

HRESULT Terminal::DebugRunVtUpgradeCorpus(TerminalVtUpgradeTestContract::Evidence* evidence) noexcept
{
    if (evidence == nullptr || evidence->sizeBytes < sizeof(*evidence))
    {
        return E_INVALIDARG;
    }
    *evidence = {};
    const VtUpgradeCorpus corpus = BuildVtUpgradeCorpus();
    evidence->inputByteCount = corpus.inputByteCount;
    if (! ComputeSha256(corpus.canonicalActionStream, evidence->corpusSha256))
    {
        return E_FAIL;
    }

    const HRESULT initializeHr = initializeEngine();
    if (FAILED(initializeHr))
    {
        return initializeHr;
    }
    _osc52Policy.store(Osc52Policy::Deny, std::memory_order_release);

    wil::unique_hwnd renderWindow(CreateWindowExW(WS_EX_TOOLWINDOW,
                                                   L"STATIC",
                                                   L"",
                                                   WS_POPUP,
                                                   0,
                                                   0,
                                                   960,
                                                   640,
                                                   nullptr,
                                                   nullptr,
                                                   g_hInstance,
                                                   nullptr));
    if (! renderWindow)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    _windowHandle.store(renderWindow.get(), std::memory_order_release);
    const auto restoreRenderWindow = wil::scope_exit([this, hwnd = renderWindow.get()]() noexcept
    {
        KillTimer(hwnd, kKittyUploadRetryTimer);
        discardDeviceResources();
        _windowHandle.store(nullptr, std::memory_order_release);
    });
    if (FAILED(ensureDeviceResources()))
    {
        return E_FAIL;
    }

    const auto placementCount = [this](size_t& count, size_t& allocatedBytes) noexcept
    {
        count = 0u;
        allocatedBytes = 0u;
        GhosttyKittyGraphics graphics = nullptr;
        if (_runtime.terminalGet(_ghosttyTerminal, GHOSTTY_TERMINAL_DATA_KITTY_GRAPHICS, &graphics) != GHOSTTY_SUCCESS ||
            graphics == nullptr)
        {
            return false;
        }
        return _runtime.kittyGraphicsGet(
                   graphics, GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_COUNT, &count) == GHOSTTY_SUCCESS &&
            _runtime.kittyGraphicsGet(
                   graphics, GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_ALLOCATED_BYTES, &allocatedBytes) == GHOSTTY_SUCCESS;
    };

    const auto captureSnapshot = [this](std::string& canonical) noexcept
    {
        canonical.clear();
        {
            std::scoped_lock lock(_terminalMutex);
            if (_runtime.renderStateBeginUpdate(_renderState, _ghosttyTerminal) != GHOSTTY_SUCCESS)
            {
                return false;
            }
        }
        // The admitted render API requires EndUpdate to complete before any
        // render-state getters or iterators are used. EndUpdate owns only the
        // render state, so terminal writers may resume before traversal.
        if (_runtime.renderStateEndUpdate(_renderState) != GHOSTTY_SUCCESS)
        {
            return false;
        }
        uint16_t columns = 0u;
        uint16_t rows = 0u;
        bool cursorVisible = false;
        bool cursorInViewport = false;
        uint16_t cursorX = 0u;
        uint16_t cursorY = 0u;
        GhosttyRenderStateCursorVisualStyle cursorStyle = GHOSTTY_RENDER_STATE_CURSOR_VISUAL_STYLE_BLOCK;
        if (_runtime.renderStateGet(_renderState, GHOSTTY_RENDER_STATE_DATA_COLS, &columns) != GHOSTTY_SUCCESS ||
            _runtime.renderStateGet(_renderState, GHOSTTY_RENDER_STATE_DATA_ROWS, &rows) != GHOSTTY_SUCCESS ||
            _runtime.renderStateGet(_renderState, GHOSTTY_RENDER_STATE_DATA_CURSOR_VISIBLE, &cursorVisible) != GHOSTTY_SUCCESS ||
            _runtime.renderStateGet(
                _renderState, GHOSTTY_RENDER_STATE_DATA_CURSOR_VIEWPORT_HAS_VALUE, &cursorInViewport) != GHOSTTY_SUCCESS ||
            _runtime.renderStateGet(
                _renderState, GHOSTTY_RENDER_STATE_DATA_CURSOR_VISUAL_STYLE, &cursorStyle) != GHOSTTY_SUCCESS ||
            (cursorInViewport &&
             (_runtime.renderStateGet(
                  _renderState, GHOSTTY_RENDER_STATE_DATA_CURSOR_VIEWPORT_X, &cursorX) != GHOSTTY_SUCCESS ||
              _runtime.renderStateGet(
                  _renderState, GHOSTTY_RENDER_STATE_DATA_CURSOR_VIEWPORT_Y, &cursorY) != GHOSTTY_SUCCESS)) ||
            _runtime.renderStateGet(
                _renderState, GHOSTTY_RENDER_STATE_DATA_ROW_ITERATOR, &_renderRows) != GHOSTTY_SUCCESS)
        {
            return false;
        }
        canonical.append(std::format("snapshot-v1;{}x{};cursor={},{},{},{},{}\n",
                                     columns,
                                     rows,
                                     cursorVisible ? 1 : 0,
                                     cursorInViewport ? 1 : 0,
                                     cursorX,
                                     cursorY,
                                     static_cast<unsigned>(cursorStyle)));
        uint16_t rowIndex = 0u;
        while (rowIndex < rows && _runtime.renderRowIteratorNext(_renderRows))
        {
            if (_runtime.renderRowGet(
                    _renderRows, GHOSTTY_RENDER_STATE_ROW_DATA_CELLS, &_renderCells) != GHOSTTY_SUCCESS)
            {
                return false;
            }
            uint16_t columnIndex = 0u;
            while (columnIndex < columns && _runtime.renderRowCellsNext(_renderCells))
            {
                GhosttyCell rawCell = 0u;
                GhosttyCellWide wide = GHOSTTY_CELL_WIDE_NARROW;
                bool hyperlink = false;
                bool selected = false;
                bool hasStyling = false;
                GhosttyStyle style{};
                style.size = sizeof(style);
                GhosttyColorRgb background{};
                GhosttyColorRgb foreground{};
                const bool rawReady = _runtime.renderRowCellsGet(
                                          _renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_RAW, &rawCell) == GHOSTTY_SUCCESS;
                const bool explicitBackground = _runtime.renderRowCellsGet(
                                                    _renderCells,
                                                    GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_BG_COLOR,
                                                    &background) == GHOSTTY_SUCCESS;
                const bool explicitForeground = _runtime.renderRowCellsGet(
                                                    _renderCells,
                                                    GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_FG_COLOR,
                                                    &foreground) == GHOSTTY_SUCCESS;
                if (! rawReady ||
                    _runtime.cellGet(rawCell, GHOSTTY_CELL_DATA_WIDE, &wide) != GHOSTTY_SUCCESS ||
                    _runtime.cellGet(rawCell, GHOSTTY_CELL_DATA_HAS_HYPERLINK, &hyperlink) != GHOSTTY_SUCCESS ||
                    _runtime.renderRowCellsGet(
                        _renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_SELECTED, &selected) != GHOSTTY_SUCCESS ||
                    _runtime.renderRowCellsGet(
                        _renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_HAS_STYLING, &hasStyling) != GHOSTTY_SUCCESS ||
                    (hasStyling &&
                     _runtime.renderRowCellsGet(
                         _renderCells, GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_STYLE, &style) != GHOSTTY_SUCCESS))
                {
                    return false;
                }
                std::array<uint8_t, 256u> graphemeBytes{};
                GhosttyBuffer grapheme{graphemeBytes.data(), graphemeBytes.size(), 0u};
                if (_runtime.renderRowCellsGet(
                        _renderCells,
                        GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_GRAPHEMES_UTF8,
                        &grapheme) != GHOSTTY_SUCCESS)
                {
                    return false;
                }
                canonical.append(std::format(
                    "cell={},{},w{},h{},s{},b{},i{},f{},k{},v{},x{},o{},u{},bg{}:{:02x}{:02x}{:02x},fg{}:{:02x}{:02x}{:02x},g{}:",
                    rowIndex,
                    columnIndex,
                    static_cast<unsigned>(wide),
                    hyperlink ? 1 : 0,
                    selected ? 1 : 0,
                    style.bold ? 1 : 0,
                    style.italic ? 1 : 0,
                    style.faint ? 1 : 0,
                    style.blink ? 1 : 0,
                    style.inverse ? 1 : 0,
                    style.invisible ? 1 : 0,
                    style.overline ? 1 : 0,
                    style.underline,
                    explicitBackground ? 1 : 0,
                    background.r,
                    background.g,
                    background.b,
                    explicitForeground ? 1 : 0,
                    foreground.r,
                    foreground.g,
                    foreground.b,
                    grapheme.len));
                canonical.append(reinterpret_cast<const char*>(grapheme.ptr), grapheme.len);
                canonical.push_back('\n');
                ++columnIndex;
            }
            if (columnIndex != columns)
            {
                return false;
            }
            ++rowIndex;
        }
        if (rowIndex != rows)
        {
            return false;
        }

        {
            std::scoped_lock lock(_terminalMutex);
            GhosttyKittyGraphics graphics = nullptr;
            size_t enginePlacementCount = 0u;
            if (_runtime.terminalGet(
                    _ghosttyTerminal, GHOSTTY_TERMINAL_DATA_KITTY_GRAPHICS, &graphics) != GHOSTTY_SUCCESS ||
                graphics == nullptr ||
                _runtime.kittyGraphicsGet(
                    graphics, GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_COUNT, &enginePlacementCount) != GHOSTTY_SUCCESS ||
                _runtime.kittyGraphicsGet(
                    graphics, GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_ITERATOR, &_kittyPlacementIterator) != GHOSTTY_SUCCESS)
            {
                return false;
            }
            // Allocation bytes are an engine storage detail, not semantic terminal
            // state. They remain separate evidence but must not perturb the snapshot
            // digest when an ABI-compatible upstream structure grows.
            canonical.append(std::format("kitty={}\n", enginePlacementCount));
            size_t observedPlacements = 0u;
            while (_runtime.kittyPlacementNext(_kittyPlacementIterator))
            {
                uint32_t imageId = 0u;
                uint32_t placementId = 0u;
                bool isVirtual = false;
                uint32_t xOffset = 0u;
                uint32_t yOffset = 0u;
                uint32_t sourceX = 0u;
                uint32_t sourceY = 0u;
                uint32_t sourceWidth = 0u;
                uint32_t sourceHeight = 0u;
                uint32_t placementColumns = 0u;
                uint32_t placementRows = 0u;
                int32_t z = 0;
                if (_runtime.kittyPlacementGet(
                        _kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_IMAGE_ID, &imageId) != GHOSTTY_SUCCESS ||
                    _runtime.kittyPlacementGet(
                        _kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_PLACEMENT_ID, &placementId) != GHOSTTY_SUCCESS ||
                    _runtime.kittyPlacementGet(
                        _kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_IS_VIRTUAL, &isVirtual) != GHOSTTY_SUCCESS ||
                    _runtime.kittyPlacementGet(
                        _kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_X_OFFSET, &xOffset) != GHOSTTY_SUCCESS ||
                    _runtime.kittyPlacementGet(
                        _kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_Y_OFFSET, &yOffset) != GHOSTTY_SUCCESS ||
                    _runtime.kittyPlacementGet(
                        _kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_SOURCE_X, &sourceX) != GHOSTTY_SUCCESS ||
                    _runtime.kittyPlacementGet(
                        _kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_SOURCE_Y, &sourceY) != GHOSTTY_SUCCESS ||
                    _runtime.kittyPlacementGet(
                        _kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_SOURCE_WIDTH, &sourceWidth) != GHOSTTY_SUCCESS ||
                    _runtime.kittyPlacementGet(
                        _kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_SOURCE_HEIGHT, &sourceHeight) != GHOSTTY_SUCCESS ||
                    _runtime.kittyPlacementGet(
                        _kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_COLUMNS, &placementColumns) != GHOSTTY_SUCCESS ||
                    _runtime.kittyPlacementGet(
                        _kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_ROWS, &placementRows) != GHOSTTY_SUCCESS ||
                    _runtime.kittyPlacementGet(
                        _kittyPlacementIterator, GHOSTTY_KITTY_GRAPHICS_PLACEMENT_DATA_Z, &z) != GHOSTTY_SUCCESS)
                {
                    return false;
                }
                GhosttyKittyGraphicsImage image = _runtime.kittyGraphicsImage(graphics, imageId);
                uint32_t width = 0u;
                uint32_t height = 0u;
                GhosttyKittyImageFormat format = GHOSTTY_KITTY_IMAGE_FORMAT_MAX_VALUE;
                const uint8_t* pixels = nullptr;
                size_t pixelBytes = 0u;
                const std::array<GhosttyKittyGraphicsImageData, 5u> imageKeys{{
                    GHOSTTY_KITTY_IMAGE_DATA_WIDTH,
                    GHOSTTY_KITTY_IMAGE_DATA_HEIGHT,
                    GHOSTTY_KITTY_IMAGE_DATA_FORMAT,
                    GHOSTTY_KITTY_IMAGE_DATA_DATA_PTR,
                    GHOSTTY_KITTY_IMAGE_DATA_DATA_LEN}};
                std::array<void*, imageKeys.size()> imageValues{{&width, &height, &format, &pixels, &pixelBytes}};
                if (image == nullptr ||
                    _runtime.kittyGraphicsImageGetMulti(
                        image, imageKeys.size(), imageKeys.data(), imageValues.data(), nullptr) != GHOSTTY_SUCCESS ||
                    (pixelBytes != 0u && pixels == nullptr))
                {
                    return false;
                }
                canonical.append(std::format("placement={},{},{},{},{},{},{},{},{},{},{},{},image={},{},{}:{}:",
                                             imageId,
                                             placementId,
                                             isVirtual ? 1 : 0,
                                             xOffset,
                                             yOffset,
                                             sourceX,
                                             sourceY,
                                             sourceWidth,
                                             sourceHeight,
                                             placementColumns,
                                             placementRows,
                                             z,
                                             width,
                                             height,
                                             static_cast<unsigned>(format),
                                             pixelBytes));
                canonical.append(reinterpret_cast<const char*>(pixels), pixelBytes);
                canonical.push_back('\n');
                ++observedPlacements;
            }
            if (observedPlacements != enginePlacementCount)
            {
                return false;
            }
        }
        return true;
    };

    constexpr std::string_view firstStyledCell = "\x1b" "c\x1b[31;44;1mA\x1b[0m";
    constexpr std::string_view replacementStyledCell = "\r\x1b[32;45;3mB\x1b[0m";
    _runtime.terminalVtWrite(
        _ghosttyTerminal, reinterpret_cast<const uint8_t*>(firstStyledCell.data()), firstStyledCell.size());
    std::string firstStyledSnapshot;
    if (! captureSnapshot(firstStyledSnapshot))
    {
        return E_FAIL;
    }
    _runtime.terminalVtWrite(_ghosttyTerminal,
                             reinterpret_cast<const uint8_t*>(replacementStyledCell.data()),
                             replacementStyledCell.size());
    std::string replacementStyledSnapshot;
    if (! captureSnapshot(replacementStyledSnapshot) || firstStyledSnapshot == replacementStyledSnapshot ||
        replacementStyledSnapshot.find("g1:B\n") == std::string::npos)
    {
        return E_FAIL;
    }

    const auto applyCorpus = [this, &corpus, &placementCount](
                                 std::string_view postlude,
                                 uint64_t& vtWriteUs,
                                 size_t& finalPlacementCount,
                                 size_t& finalPlacementBytes) noexcept
    {
        _kittyPipeline.Invalidate(++_kittyRequestId);
        _kittyPlacements.clear();
        _kittyImages.clear();
        _kittyCachedBytes = 0u;
        _kittyTargetStorageGeneration = 0u;
        DebugClearCapturedInput();
        uint64_t elapsedNanoseconds = 0u;
        const auto write = [this, &elapsedNanoseconds](std::string_view bytes) noexcept
        {
            const auto startedAt = std::chrono::steady_clock::now();
            _runtime.terminalVtWrite(
                _ghosttyTerminal, reinterpret_cast<const uint8_t*>(bytes.data()), bytes.size());
            elapsedNanoseconds += static_cast<uint64_t>(std::max<int64_t>(
                0,
                std::chrono::duration_cast<std::chrono::nanoseconds>(
                    std::chrono::steady_clock::now() - startedAt).count()));
        };
        write(corpus.prelude);
        _columns = 96u;
        _rows = 32u;
        if (_runtime.terminalResize(_ghosttyTerminal, _columns, _rows, 8u, 16u) != GHOSTTY_SUCCESS)
        {
            return false;
        }
        write(corpus.kittyPartial);
        size_t pendingCount = 0u;
        size_t pendingBytes = 0u;
        if (! placementCount(pendingCount, pendingBytes) || pendingCount != 0u)
        {
            return false;
        }
        write(corpus.kittyCompletion);
        write(postlude);
        _columns = 120u;
        _rows = 40u;
        if (_runtime.terminalResize(_ghosttyTerminal, _columns, _rows, 8u, 16u) != GHOSTTY_SUCCESS ||
            ! placementCount(finalPlacementCount, finalPlacementBytes) || finalPlacementCount != 1u)
        {
            return false;
        }
        vtWriteUs = std::max<uint64_t>(1u, (elapsedNanoseconds + 999u) / 1'000u);
        return true;
    };

    constexpr uint32_t warmupCount = 5u;
    std::array<char, TerminalVtUpgradeTestContract::kSha256TextBytes> expectedOutputDigest{};
    std::array<char, TerminalVtUpgradeTestContract::kSha256TextBytes> expectedSnapshotDigest{};
    bool expectedDigestsSet = false;
    for (uint32_t ordinal = 0u; ordinal < warmupCount + evidence->sampleCount; ++ordinal)
    {
        uint64_t vtWriteUs = 0u;
        size_t enginePlacementCount = 0u;
        size_t enginePlacementBytes = 0u;
        if (! applyCorpus(corpus.postlude, vtWriteUs, enginePlacementCount, enginePlacementBytes))
        {
            return E_FAIL;
        }
        ++evidence->kittyPendingProbeCount;
        ++evidence->kittyCompletedProbeCount;

        InvalidateRect(renderWindow.get(), nullptr, FALSE);
        render();
        const uint64_t requestId = _kittyRequestId;
        const auto kittyDeadline = std::chrono::steady_clock::now() + std::chrono::seconds(2);
        bool kittyReady = false;
        while (std::chrono::steady_clock::now() < kittyDeadline)
        {
            consumeKittyReady();
            for (const auto& [imageKey, image] : _kittyImages)
            {
                if (imageKey.imageId == 77u && image.state == KittyImageState::Ready &&
                    image.pendingRequestId == 0u)
                {
                    kittyReady = true;
                    break;
                }
            }
            if (kittyReady || _kittyPipeline.GetStats().latestRequestId > requestId)
            {
                break;
            }
            std::this_thread::yield();
        }
        if (! kittyReady)
        {
            return E_FAIL;
        }
        InvalidateRect(renderWindow.get(), nullptr, FALSE);
        render();

        std::string snapshotCanonical;
        const auto snapshotStartedAt = std::chrono::steady_clock::now();
        if (! captureSnapshot(snapshotCanonical))
        {
            return E_FAIL;
        }
        const uint64_t snapshotUs = ElapsedUs(snapshotStartedAt);
        std::array<char, TerminalVtUpgradeTestContract::kSha256TextBytes> snapshotDigest{};
        if (! ComputeSha256(snapshotCanonical, snapshotDigest))
        {
            return E_FAIL;
        }
        const std::string ptyOutput = DebugTakeCapturedInput();
        std::array<char, TerminalVtUpgradeTestContract::kSha256TextBytes> outputDigest{};
        if (! ComputeSha256(ptyOutput, outputDigest))
        {
            return E_FAIL;
        }

        InvalidateRect(renderWindow.get(), nullptr, FALSE);
        const auto hostedRenderStartedAt = std::chrono::steady_clock::now();
        render();
        const uint64_t hostedRenderUs = ElapsedUs(hostedRenderStartedAt);
        if (! _renderTarget)
        {
            return E_FAIL;
        }

        if (! expectedDigestsSet)
        {
            expectedOutputDigest = outputDigest;
            expectedSnapshotDigest = snapshotDigest;
            expectedDigestsSet = true;
        }
        else if (outputDigest != expectedOutputDigest || snapshotDigest != expectedSnapshotDigest)
        {
            return E_FAIL;
        }

        const TerminalKittyPipelineStats kittyStats = _kittyPipeline.GetStats();
        evidence->kittyPlacementCount = enginePlacementCount;
        evidence->kittyPlacementBytes = enginePlacementBytes;
        evidence->kittyQueuedBytes = kittyStats.queuedBytes;
        evidence->kittyActiveSourceBytes = kittyStats.activeSourceBytes;
        evidence->kittyActiveConvertedBytes = kittyStats.activeConvertedBytes;
        evidence->kittyReadyBytes = kittyStats.readyBytes;
        evidence->kittyResidentBytes = kittyStats.residentBytes;
        evidence->kittyMemoryHighWaterBytes = std::max<uint64_t>(
            evidence->kittyMemoryHighWaterBytes, kittyStats.memoryHighWaterBytes);
        evidence->kittyCacheBytes = _kittyCachedBytes;
        evidence->kittyPinnedBytes = _kittyCapturePinnedBytes;
        evidence->kittyFrameUploadBytes = _kittyFrameUploadBytes;
        evidence->kittyFrameUploadCount = _kittyFrameUploadCount;
        ObserveProcessMemory(*evidence);

        if (ordinal >= warmupCount)
        {
            const size_t sample = ordinal - warmupCount;
            evidence->vtWriteSamplesUs[sample] = vtWriteUs;
            evidence->snapshotSamplesUs[sample] = snapshotUs;
            evidence->hostedRenderSamplesUs[sample] = hostedRenderUs;
        }
    }

    uint64_t controlVtWriteUs = 0u;
    size_t controlPlacementCount = 0u;
    size_t controlPlacementBytes = 0u;
    if (! applyCorpus(corpus.postludeWithoutExpectedDelta,
                      controlVtWriteUs,
                      controlPlacementCount,
                      controlPlacementBytes) ||
        controlVtWriteUs == 0u || controlPlacementCount != 1u || controlPlacementBytes == 0u)
    {
        return E_FAIL;
    }
    InvalidateRect(renderWindow.get(), nullptr, FALSE);
    render();
    const uint64_t controlRequestId = _kittyRequestId;
    const auto controlKittyDeadline = std::chrono::steady_clock::now() + std::chrono::seconds(2);
    bool controlKittyReady = false;
    while (std::chrono::steady_clock::now() < controlKittyDeadline)
    {
        consumeKittyReady();
        for (const auto& [imageKey, image] : _kittyImages)
        {
            if (imageKey.imageId == 77u && image.state == KittyImageState::Ready &&
                image.pendingRequestId == 0u)
            {
                controlKittyReady = true;
                break;
            }
        }
        if (controlKittyReady || _kittyPipeline.GetStats().latestRequestId > controlRequestId)
        {
            break;
        }
        std::this_thread::yield();
    }
    std::string controlSnapshotCanonical;
    if (! controlKittyReady || ! captureSnapshot(controlSnapshotCanonical) ||
        ! ComputeSha256(controlSnapshotCanonical, evidence->snapshotWithoutExpectedDeltaSha256))
    {
        return E_FAIL;
    }
    const std::string controlPtyOutput = DebugTakeCapturedInput();
    if (! ComputeSha256(controlPtyOutput, evidence->ptyOutputWithoutExpectedDeltaSha256))
    {
        return E_FAIL;
    }

    evidence->ptyOutputSha256 = expectedOutputDigest;
    evidence->snapshotSha256 = expectedSnapshotDigest;
    evidence->assertionCount = 14u;
    return evidence->kittyPendingProbeCount == warmupCount + evidence->sampleCount &&
            evidence->kittyCompletedProbeCount == warmupCount + evidence->sampleCount &&
            evidence->kittyPlacementCount == 1u && evidence->kittyPlacementBytes != 0u &&
            evidence->kittyPlacementBytes <= kMaximumKittyPlacementBytes &&
            evidence->kittyCacheBytes == 4u && evidence->kittyPinnedBytes == 4u &&
            evidence->peakPrivateBytes != 0u && evidence->peakWorkingSetBytes != 0u
        ? S_OK
        : E_FAIL;
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderTerminalDebugVtUpgradeCorpus(
    TerminalVtUpgradeTestContract::Evidence* evidence) noexcept
{
    Terminal terminal;
    return terminal.DebugRunVtUpgradeCorpus(evidence);
}

#endif

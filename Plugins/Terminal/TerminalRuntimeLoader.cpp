#include "TerminalRuntimeLoader.h"

#include <red_salamander/terminal_runtime_identity.generated.h>

#include <array>
#include <atomic>
#include <bcrypt.h>
#include <cstddef>
#include <cstdint>
#include <format>
#include <limits>
#include <mutex>
#include <span>
#include <string>
#include <string_view>
#include <type_traits>
#include <vector>

#include <wincodec.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/com.h>
#pragma warning(pop)

#include "Helpers.h"
#include "resource.h"

#pragma comment(lib, "windowscodecs")
#pragma comment(lib, "bcrypt")

extern HINSTANCE g_hInstance;

namespace
{
thread_local IWICImagingFactory* g_pngDecodeWicFactory = nullptr;
template <typename T>
concept HasSizeMember = requires(T value) { value.size; };

static_assert(std::is_standard_layout_v<GhosttyTerminalModeConfig>);
static_assert(! HasSizeMember<GhosttyTerminalModeConfig>);
static_assert(offsetof(GhosttyTerminalModeConfig, value) == sizeof(GhosttyMode));
} // namespace

TerminalPngDecodeThreadContext::TerminalPngDecodeThreadContext() noexcept
{
    const HRESULT coInitializeResult = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (SUCCEEDED(coInitializeResult))
    {
        _coUninitialize.activate();
    }
    if (FAILED(coInitializeResult) && coInitializeResult != RPC_E_CHANGED_MODE)
    {
        _status = coInitializeResult;
        return;
    }

    HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory2, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(_factory.put()));
    if (FAILED(hr))
    {
        hr = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(_factory.put()));
    }
    if (FAILED(hr) || ! _factory)
    {
        _status = FAILED(hr) ? hr : E_NOINTERFACE;
        return;
    }

    _previousFactory      = g_pngDecodeWicFactory;
    g_pngDecodeWicFactory = _factory.get();
    _installed            = true;
    _status               = S_OK;
}

TerminalPngDecodeThreadContext::~TerminalPngDecodeThreadContext() noexcept
{
    if (_installed)
    {
        g_pngDecodeWicFactory = _previousFactory;
    }
}

HRESULT TerminalPngDecodeThreadContext::Status() const noexcept
{
    return _status;
}

namespace
{
constexpr uint64_t kMaximumKittyPixels             = 16u * 1024u * 1024u;
constexpr size_t kMaximumKittyDecodedBytes         = 64u * 1024u * 1024u;
constexpr uint64_t kExpectedRuntimeBytes           = RSTerminalRuntimeIdentity::kSizeBytes;
constexpr std::wstring_view kExpectedRuntimeSha256 = RSTerminalRuntimeIdentity::kSha256;
static_assert(kExpectedRuntimeBytes > 0u);
static_assert(kExpectedRuntimeSha256.size() == 64u);

#if defined(_M_X64)
constexpr WORD kExpectedRuntimeMachine = IMAGE_FILE_MACHINE_AMD64;
#elif defined(_M_ARM64)
constexpr WORD kExpectedRuntimeMachine = IMAGE_FILE_MACHINE_ARM64;
#else
#error Terminal private Ghostty runtime identity is defined only for x64 and ARM64.
#endif

std::mutex g_systemCallbackMutex;
uint32_t g_systemCallbackUsers = 0u;
std::atomic<decltype(&ghostty_alloc)> g_ghosttyAlloc{nullptr};
std::atomic<decltype(&ghostty_free)> g_ghosttyFree{nullptr};

[[nodiscard]] bool DecodePng(void* /*userData*/, const GhosttyAllocator* allocator, const uint8_t* data, size_t dataLength, GhosttySysImage* output) noexcept
{
    if (data == nullptr || output == nullptr || dataLength == 0u || dataLength > static_cast<size_t>(MAXDWORD))
    {
        return false;
    }
    *output = {};

    IWICImagingFactory* const factory = g_pngDecodeWicFactory;
    if (factory == nullptr)
    {
        return false;
    }

    wil::com_ptr<IWICStream> stream;
    if (FAILED(factory->CreateStream(stream.put())) || ! stream ||
        FAILED(stream->InitializeFromMemory(const_cast<BYTE*>(data), static_cast<DWORD>(dataLength))))
    {
        return false;
    }

    wil::com_ptr<IWICBitmapDecoder> decoder;
    if (FAILED(factory->CreateDecoderFromStream(stream.get(), &GUID_ContainerFormatPng, WICDecodeMetadataCacheOnDemand, decoder.put())) || ! decoder)
    {
        return false;
    }

    wil::com_ptr<IWICBitmapFrameDecode> frame;
    if (FAILED(decoder->GetFrame(0u, frame.put())) || ! frame)
    {
        return false;
    }
    UINT width  = 0u;
    UINT height = 0u;
    if (FAILED(frame->GetSize(&width, &height)) || width == 0u || height == 0u)
    {
        return false;
    }
    const uint64_t pixels       = static_cast<uint64_t>(width) * static_cast<uint64_t>(height);
    const uint64_t decodedBytes = pixels * 4u;
    if (pixels > kMaximumKittyPixels || decodedBytes > kMaximumKittyDecodedBytes || decodedBytes > static_cast<uint64_t>((std::numeric_limits<UINT>::max)()))
    {
        return false;
    }

    wil::com_ptr<IWICFormatConverter> converter;
    if (FAILED(factory->CreateFormatConverter(converter.put())) || ! converter ||
        FAILED(converter->Initialize(frame.get(), GUID_WICPixelFormat32bppRGBA, WICBitmapDitherTypeNone, nullptr, 0.0, WICBitmapPaletteTypeCustom)))
    {
        return false;
    }

    const auto allocate = g_ghosttyAlloc.load(std::memory_order_acquire);
    const auto release  = g_ghosttyFree.load(std::memory_order_acquire);
    if (allocate == nullptr || release == nullptr)
    {
        return false;
    }
    const size_t byteCount = static_cast<size_t>(decodedBytes);
    uint8_t* decoded       = allocate(allocator, byteCount);
    if (decoded == nullptr)
    {
        return false;
    }
    auto releaseOnFailure = wil::scope_exit([&]() noexcept { release(allocator, decoded, byteCount); });
    const UINT stride     = width * 4u;
    if (FAILED(converter->CopyPixels(nullptr, stride, static_cast<UINT>(byteCount), decoded)))
    {
        return false;
    }

    output->width    = width;
    output->height   = height;
    output->data     = decoded;
    output->data_len = byteCount;
    releaseOnFailure.release();
    return true;
}

[[nodiscard]] std::wstring GetRuntimePath() noexcept
{
    std::array<wchar_t, 32768> modulePath{};
    const DWORD copied = GetModuleFileNameW(g_hInstance, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    if (copied == 0u || copied >= modulePath.size())
    {
        return {};
    }

    std::wstring path(modulePath.data(), copied);
    const size_t slash = path.find_last_of(L"\\/");
    if (slash == std::wstring::npos)
    {
        return {};
    }
    path.resize(slash + 1u);
    path.append(L"TerminalRuntime\\ghostty-vt.dll");
    return path;
}

[[nodiscard]] HRESULT ReadRuntimeBytes(HANDLE file, uint64_t offset, std::span<std::byte> destination) noexcept
{
    if (file == nullptr || file == INVALID_HANDLE_VALUE || destination.empty() ||
        destination.size() > static_cast<size_t>((std::numeric_limits<DWORD>::max)()) || offset > static_cast<uint64_t>((std::numeric_limits<LONGLONG>::max)()))
    {
        return E_INVALIDARG;
    }

    LARGE_INTEGER position{};
    position.QuadPart = static_cast<LONGLONG>(offset);
    if (SetFilePointerEx(file, position, nullptr, FILE_BEGIN) == FALSE)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    DWORD bytesRead = 0u;
    if (ReadFile(file, destination.data(), static_cast<DWORD>(destination.size()), &bytesRead, nullptr) == FALSE)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    return static_cast<size_t>(bytesRead) == destination.size() ? S_OK : HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
}

template <typename T> [[nodiscard]] HRESULT ReadRuntimeValue(HANDLE file, uint64_t offset, T& value) noexcept
{
    return ReadRuntimeBytes(file, offset, std::as_writable_bytes(std::span<T>(&value, 1u)));
}

[[nodiscard]] HRESULT ComputeRuntimeSha256(HANDLE file, std::wstring& digestText) noexcept
{
    digestText.clear();
    LARGE_INTEGER beginning{};
    if (SetFilePointerEx(file, beginning, nullptr, FILE_BEGIN) == FALSE)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    wil::unique_bcrypt_algorithm algorithm;
    NTSTATUS status = BCryptOpenAlgorithmProvider(algorithm.put(), BCRYPT_SHA256_ALGORITHM, nullptr, 0u);
    if (! BCRYPT_SUCCESS(status))
    {
        return HRESULT_FROM_NT(status);
    }

    DWORD hashObjectBytes = 0u;
    DWORD propertyBytes   = 0u;
    status = BCryptGetProperty(algorithm.get(), BCRYPT_OBJECT_LENGTH, reinterpret_cast<PUCHAR>(&hashObjectBytes), sizeof(hashObjectBytes), &propertyBytes, 0u);
    if (! BCRYPT_SUCCESS(status) || hashObjectBytes == 0u)
    {
        return BCRYPT_SUCCESS(status) ? E_FAIL : HRESULT_FROM_NT(status);
    }

    std::vector<std::byte> hashObject(hashObjectBytes);
    wil::unique_bcrypt_hash hash;
    status = BCryptCreateHash(algorithm.get(), hash.put(), reinterpret_cast<PUCHAR>(hashObject.data()), hashObjectBytes, nullptr, 0u, 0u);
    if (! BCRYPT_SUCCESS(status))
    {
        return HRESULT_FROM_NT(status);
    }

    std::array<uint8_t, 64u * 1024u> buffer{};
    for (;;)
    {
        DWORD bytesRead = 0u;
        if (ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &bytesRead, nullptr) == FALSE)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (bytesRead == 0u)
        {
            break;
        }
        status = BCryptHashData(hash.get(), buffer.data(), bytesRead, 0u);
        if (! BCRYPT_SUCCESS(status))
        {
            return HRESULT_FROM_NT(status);
        }
    }

    std::array<uint8_t, 32u> digest{};
    status = BCryptFinishHash(hash.get(), digest.data(), static_cast<ULONG>(digest.size()), 0u);
    if (! BCRYPT_SUCCESS(status))
    {
        return HRESULT_FROM_NT(status);
    }

    constexpr wchar_t kHex[] = L"0123456789abcdef";
    digestText.resize(digest.size() * 2u);
    for (size_t index = 0u; index < digest.size(); ++index)
    {
        digestText[index * 2u]        = kHex[digest[index] >> 4u];
        digestText[(index * 2u) + 1u] = kHex[digest[index] & 0x0Fu];
    }
    return S_OK;
}

[[nodiscard]] HRESULT ValidateRuntimeIdentity(HANDLE file, std::wstring& errorText) noexcept
{
    BY_HANDLE_FILE_INFORMATION information{};
    if (GetFileInformationByHandle(file, &information) == FALSE)
    {
        const DWORD error = GetLastError();
        errorText         = FormatStringResource(g_hInstance, IDS_TERMINAL_RUNTIME_METADATA_ERROR, error);
        return HRESULT_FROM_WIN32(error);
    }
    if ((information.dwFileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) != 0u)
    {
        errorText = LoadStringResource(g_hInstance, IDS_TERMINAL_RUNTIME_REGULAR_FILE_ERROR);
        return HRESULT_FROM_WIN32(ERROR_REPARSE_TAG_INVALID);
    }

    const uint64_t fileBytes = (static_cast<uint64_t>(information.nFileSizeHigh) << 32u) | information.nFileSizeLow;
    if (fileBytes != kExpectedRuntimeBytes)
    {
        errorText = FormatStringResource(g_hInstance, IDS_TERMINAL_RUNTIME_SIZE_ERROR, kExpectedRuntimeBytes, fileBytes);
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }

    IMAGE_DOS_HEADER dosHeader{};
    HRESULT hr = ReadRuntimeValue(file, 0u, dosHeader);
    if (FAILED(hr) || dosHeader.e_magic != IMAGE_DOS_SIGNATURE || dosHeader.e_lfanew < static_cast<LONG>(sizeof(IMAGE_DOS_HEADER)))
    {
        errorText = LoadStringResource(g_hInstance, IDS_TERMINAL_RUNTIME_DOS_HEADER_ERROR);
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_BAD_EXE_FORMAT);
    }
    const uint64_t ntOffset = static_cast<uint64_t>(dosHeader.e_lfanew);
    if (ntOffset > fileBytes - sizeof(DWORD) - sizeof(IMAGE_FILE_HEADER))
    {
        errorText = LoadStringResource(g_hInstance, IDS_TERMINAL_RUNTIME_PE_HEADER_ERROR);
        return HRESULT_FROM_WIN32(ERROR_BAD_EXE_FORMAT);
    }

    DWORD signature = 0u;
    IMAGE_FILE_HEADER fileHeader{};
    hr = ReadRuntimeValue(file, ntOffset, signature);
    if (SUCCEEDED(hr))
    {
        hr = ReadRuntimeValue(file, ntOffset + sizeof(signature), fileHeader);
    }
    if (FAILED(hr) || signature != IMAGE_NT_SIGNATURE || fileHeader.Machine != kExpectedRuntimeMachine)
    {
        errorText = LoadStringResource(g_hInstance, IDS_TERMINAL_RUNTIME_MACHINE_ERROR);
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_BAD_EXE_FORMAT);
    }

    std::wstring digestText;
    hr = ComputeRuntimeSha256(file, digestText);
    if (FAILED(hr))
    {
        errorText = FormatStringResource(g_hInstance, IDS_TERMINAL_RUNTIME_SHA256_ERROR, static_cast<uint32_t>(hr));
        return hr;
    }
    if (digestText != kExpectedRuntimeSha256)
    {
        errorText = FormatStringResource(g_hInstance, IDS_TERMINAL_RUNTIME_SHA256_MISMATCH, digestText);
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }
    return S_OK;
}

[[nodiscard]] HRESULT GetLockedRuntimePath(HANDLE file, std::wstring& path, std::wstring& errorText) noexcept
{
    std::array<wchar_t, 32768u> resolvedPath{};
    const DWORD copied = GetFinalPathNameByHandleW(file, resolvedPath.data(), static_cast<DWORD>(resolvedPath.size()), FILE_NAME_NORMALIZED | VOLUME_NAME_DOS);
    if (copied == 0u || copied >= resolvedPath.size())
    {
        const DWORD error = copied == 0u ? GetLastError() : ERROR_FILENAME_EXCED_RANGE;
        errorText         = FormatStringResource(g_hInstance, IDS_TERMINAL_RUNTIME_LOCK_PATH_ERROR, error);
        return HRESULT_FROM_WIN32(error);
    }
    path.assign(resolvedPath.data(), copied);
    return S_OK;
}

template <typename T> [[nodiscard]] bool ResolveFunction(HMODULE module, const char* name, T& output) noexcept
{
#pragma warning(push)
#pragma warning(disable : 4191) // Closed export table validated before the cast.
    output = reinterpret_cast<T>(GetProcAddress(module, name));
#pragma warning(pop)
    return output != nullptr;
}
} // namespace

HRESULT TerminalRuntimeLoader::Load() noexcept
{
    if (_module)
    {
        return S_OK;
    }

    _errorText.clear();
    const std::wstring runtimePath = GetRuntimePath();
    if (runtimePath.empty())
    {
        _errorText = LoadStringResource(g_hInstance, IDS_TERMINAL_PLUGIN_PATH_ERROR);
        return E_FAIL;
    }

    wil::unique_hfile runtimeFile(CreateFileW(runtimePath.c_str(),
                                              GENERIC_READ,
                                              FILE_SHARE_READ,
                                              nullptr,
                                              OPEN_EXISTING,
                                              FILE_ATTRIBUTE_NORMAL | FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_SEQUENTIAL_SCAN,
                                              nullptr));
    if (! runtimeFile)
    {
        const DWORD error = GetLastError();
        _errorText        = FormatStringResource(g_hInstance, IDS_TERMINAL_RUNTIME_OPEN_ERROR, error, runtimePath);
        return HRESULT_FROM_WIN32(error);
    }

    HRESULT identityHr = ValidateRuntimeIdentity(runtimeFile.get(), _errorText);
    if (FAILED(identityHr))
    {
        return identityHr;
    }

    std::wstring lockedRuntimePath;
    HRESULT lockedPathHr = GetLockedRuntimePath(runtimeFile.get(), lockedRuntimePath, _errorText);
    if (FAILED(lockedPathHr))
    {
        return lockedPathHr;
    }

    wil::unique_hmodule module(LoadLibraryExW(lockedRuntimePath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_SYSTEM32));
    if (! module)
    {
        const DWORD error = GetLastError();
        _errorText        = FormatStringResource(g_hInstance, IDS_TERMINAL_RUNTIME_LOAD_ERROR, error);
        return HRESULT_FROM_WIN32(error);
    }

    const bool complete =
        ResolveFunction(module.get(), "ghostty_build_info", buildInfo) && ResolveFunction(module.get(), "ghostty_terminal_new", terminalNew) &&
        ResolveFunction(module.get(), "ghostty_terminal_free", terminalFree) && ResolveFunction(module.get(), "ghostty_terminal_resize", terminalResize) &&
        ResolveFunction(module.get(), "ghostty_terminal_set", terminalSet) && ResolveFunction(module.get(), "ghostty_terminal_get", terminalGet) &&
        ResolveFunction(module.get(), "ghostty_terminal_grid_ref", terminalGridRef) &&
        ResolveFunction(module.get(), "ghostty_grid_ref_hyperlink_uri", gridRefHyperlinkUri) &&
        ResolveFunction(module.get(), "ghostty_terminal_grid_ref_track", terminalGridRefTrack) &&
        ResolveFunction(module.get(), "ghostty_tracked_grid_ref_free", trackedGridRefFree) &&
        ResolveFunction(module.get(), "ghostty_tracked_grid_ref_snapshot", trackedGridRefSnapshot) &&
        ResolveFunction(module.get(), "ghostty_terminal_vt_write", terminalVtWrite) &&
        ResolveFunction(module.get(), "ghostty_terminal_select_word", terminalSelectWord) &&
        ResolveFunction(module.get(), "ghostty_terminal_select_all", terminalSelectAll) &&
        ResolveFunction(module.get(), "ghostty_terminal_selection_format_buf", terminalSelectionFormatBuffer) &&
        ResolveFunction(module.get(), "ghostty_terminal_scroll_viewport", terminalScrollViewport) &&
        ResolveFunction(module.get(), "ghostty_paste_is_safe", pasteIsSafe) && ResolveFunction(module.get(), "ghostty_paste_encode", pasteEncode) &&
        ResolveFunction(module.get(), "ghostty_render_state_new", renderStateNew) &&
        ResolveFunction(module.get(), "ghostty_render_state_free", renderStateFree) &&
        ResolveFunction(module.get(), "ghostty_render_state_begin_update", renderStateBeginUpdate) &&
        ResolveFunction(module.get(), "ghostty_render_state_end_update", renderStateEndUpdate) &&
        ResolveFunction(module.get(), "ghostty_render_state_get", renderStateGet) &&
        ResolveFunction(module.get(), "ghostty_render_state_set", renderStateSet) &&
        ResolveFunction(module.get(), "ghostty_render_state_row_iterator_new", renderRowIteratorNew) &&
        ResolveFunction(module.get(), "ghostty_render_state_row_iterator_free", renderRowIteratorFree) &&
        ResolveFunction(module.get(), "ghostty_render_state_row_iterator_next", renderRowIteratorNext) &&
        ResolveFunction(module.get(), "ghostty_render_state_row_get", renderRowGet) &&
        ResolveFunction(module.get(), "ghostty_render_state_row_set", renderRowSet) &&
        ResolveFunction(module.get(), "ghostty_render_state_row_cells_new", renderRowCellsNew) &&
        ResolveFunction(module.get(), "ghostty_render_state_row_cells_free", renderRowCellsFree) &&
        ResolveFunction(module.get(), "ghostty_render_state_row_cells_next", renderRowCellsNext) &&
        ResolveFunction(module.get(), "ghostty_render_state_row_cells_get", renderRowCellsGet) && ResolveFunction(module.get(), "ghostty_cell_get", cellGet) &&
        ResolveFunction(module.get(), "ghostty_key_encoder_new", keyEncoderNew) && ResolveFunction(module.get(), "ghostty_key_encoder_free", keyEncoderFree) &&
        ResolveFunction(module.get(), "ghostty_key_encoder_setopt_from_terminal", keyEncoderSetFromTerminal) &&
        ResolveFunction(module.get(), "ghostty_key_encoder_encode", keyEncoderEncode) && ResolveFunction(module.get(), "ghostty_key_event_new", keyEventNew) &&
        ResolveFunction(module.get(), "ghostty_key_event_free", keyEventFree) &&
        ResolveFunction(module.get(), "ghostty_key_event_set_action", keyEventSetAction) &&
        ResolveFunction(module.get(), "ghostty_key_event_set_key", keyEventSetKey) &&
        ResolveFunction(module.get(), "ghostty_key_event_set_mods", keyEventSetMods) &&
        ResolveFunction(module.get(), "ghostty_key_event_set_utf8", keyEventSetUtf8) &&
        ResolveFunction(module.get(), "ghostty_key_event_set_unshifted_codepoint", keyEventSetUnshiftedCodepoint) &&
        ResolveFunction(module.get(), "ghostty_mouse_encoder_new", mouseEncoderNew) &&
        ResolveFunction(module.get(), "ghostty_mouse_encoder_free", mouseEncoderFree) &&
        ResolveFunction(module.get(), "ghostty_mouse_encoder_setopt", mouseEncoderSetOption) &&
        ResolveFunction(module.get(), "ghostty_mouse_encoder_setopt_from_terminal", mouseEncoderSetFromTerminal) &&
        ResolveFunction(module.get(), "ghostty_mouse_encoder_encode", mouseEncoderEncode) &&
        ResolveFunction(module.get(), "ghostty_mouse_event_new", mouseEventNew) && ResolveFunction(module.get(), "ghostty_mouse_event_free", mouseEventFree) &&
        ResolveFunction(module.get(), "ghostty_mouse_event_set_action", mouseEventSetAction) &&
        ResolveFunction(module.get(), "ghostty_mouse_event_set_button", mouseEventSetButton) &&
        ResolveFunction(module.get(), "ghostty_mouse_event_clear_button", mouseEventClearButton) &&
        ResolveFunction(module.get(), "ghostty_mouse_event_set_mods", mouseEventSetMods) &&
        ResolveFunction(module.get(), "ghostty_mouse_event_set_position", mouseEventSetPosition) && ResolveFunction(module.get(), "ghostty_alloc", alloc) &&
        ResolveFunction(module.get(), "ghostty_free", free) && ResolveFunction(module.get(), "ghostty_sys_set", sysSet) &&
        ResolveFunction(module.get(), "ghostty_kitty_graphics_get", kittyGraphicsGet) &&
        ResolveFunction(module.get(), "ghostty_kitty_graphics_image", kittyGraphicsImage) &&
        ResolveFunction(module.get(), "ghostty_kitty_graphics_image_get_multi", kittyGraphicsImageGetMulti) &&
        ResolveFunction(module.get(), "ghostty_kitty_graphics_placement_iterator_new", kittyPlacementIteratorNew) &&
        ResolveFunction(module.get(), "ghostty_kitty_graphics_placement_iterator_free", kittyPlacementIteratorFree) &&
        ResolveFunction(module.get(), "ghostty_kitty_graphics_placement_next", kittyPlacementNext) &&
        ResolveFunction(module.get(), "ghostty_kitty_graphics_placement_get", kittyPlacementGet) &&
        ResolveFunction(module.get(), "ghostty_kitty_graphics_placement_render_info", kittyPlacementRenderInfo) &&
        ResolveFunction(module.get(), "ghostty_formatter_terminal_new", formatterTerminalNew) &&
        ResolveFunction(module.get(), "ghostty_formatter_format_buf", formatterFormatBuffer) &&
        ResolveFunction(module.get(), "ghostty_formatter_free", formatterFree);
    if (! complete)
    {
        Unload();
        _errorText = LoadStringResource(g_hInstance, IDS_TERMINAL_RUNTIME_API_ERROR);
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }

    bool simd                    = true;
    bool kittyGraphics           = false;
    GhosttyOptimizeMode optimize = GHOSTTY_OPTIMIZE_DEBUG;
    if (buildInfo(GHOSTTY_BUILD_INFO_SIMD, &simd) != GHOSTTY_SUCCESS || buildInfo(GHOSTTY_BUILD_INFO_KITTY_GRAPHICS, &kittyGraphics) != GHOSTTY_SUCCESS ||
        buildInfo(GHOSTTY_BUILD_INFO_OPTIMIZE, &optimize) != GHOSTTY_SUCCESS || simd || ! kittyGraphics || optimize != GHOSTTY_OPTIMIZE_RELEASE_FAST)
    {
        Unload();
        _errorText = LoadStringResource(g_hInstance, IDS_TERMINAL_RUNTIME_BUILD_OPTIONS_ERROR);
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }

    {
        std::scoped_lock lock(g_systemCallbackMutex);
        if (g_systemCallbackUsers == 0u)
        {
            g_ghosttyAlloc.store(alloc, std::memory_order_release);
            g_ghosttyFree.store(free, std::memory_order_release);
            const GhosttySysDecodePngFn decoder = &DecodePng;
            if (sysSet(GHOSTTY_SYS_OPT_DECODE_PNG, reinterpret_cast<const void*>(decoder)) != GHOSTTY_SUCCESS)
            {
                g_ghosttyFree.store(nullptr, std::memory_order_release);
                g_ghosttyAlloc.store(nullptr, std::memory_order_release);
                Unload();
                _errorText = LoadStringResource(g_hInstance, IDS_TERMINAL_RUNTIME_PNG_DECODER_ERROR);
                return E_FAIL;
            }
        }
        ++g_systemCallbackUsers;
        _systemCallbacksAcquired = true;
    }

    _runtimeFile = std::move(runtimeFile);
    _module      = std::move(module);
    return S_OK;
}

void TerminalRuntimeLoader::Unload() noexcept
{
    if (_systemCallbacksAcquired)
    {
        std::scoped_lock lock(g_systemCallbackMutex);
        if (g_systemCallbackUsers != 0u)
        {
            --g_systemCallbackUsers;
        }
        if (g_systemCallbackUsers == 0u && sysSet != nullptr)
        {
            static_cast<void>(sysSet(GHOSTTY_SYS_OPT_DECODE_PNG, nullptr));
            g_ghosttyFree.store(nullptr, std::memory_order_release);
            g_ghosttyAlloc.store(nullptr, std::memory_order_release);
        }
        _systemCallbacksAcquired = false;
    }
    kittyPlacementRenderInfo      = nullptr;
    kittyPlacementGet             = nullptr;
    kittyPlacementNext            = nullptr;
    kittyPlacementIteratorFree    = nullptr;
    kittyPlacementIteratorNew     = nullptr;
    kittyGraphicsImageGetMulti    = nullptr;
    kittyGraphicsImage            = nullptr;
    kittyGraphicsGet              = nullptr;
    sysSet                        = nullptr;
    free                          = nullptr;
    alloc                         = nullptr;
    mouseEventSetPosition         = nullptr;
    mouseEventSetMods             = nullptr;
    mouseEventClearButton         = nullptr;
    mouseEventSetButton           = nullptr;
    mouseEventSetAction           = nullptr;
    mouseEventFree                = nullptr;
    mouseEventNew                 = nullptr;
    mouseEncoderEncode            = nullptr;
    mouseEncoderSetFromTerminal   = nullptr;
    mouseEncoderSetOption         = nullptr;
    mouseEncoderFree              = nullptr;
    mouseEncoderNew               = nullptr;
    keyEventSetUnshiftedCodepoint = nullptr;
    keyEventSetUtf8               = nullptr;
    keyEventSetMods               = nullptr;
    keyEventSetKey                = nullptr;
    keyEventSetAction             = nullptr;
    keyEventFree                  = nullptr;
    keyEventNew                   = nullptr;
    keyEncoderEncode              = nullptr;
    keyEncoderSetFromTerminal     = nullptr;
    keyEncoderFree                = nullptr;
    keyEncoderNew                 = nullptr;
    cellGet                       = nullptr;
    renderRowCellsGet             = nullptr;
    renderRowCellsNext            = nullptr;
    renderRowCellsFree            = nullptr;
    renderRowCellsNew             = nullptr;
    renderRowSet                  = nullptr;
    renderRowGet                  = nullptr;
    renderRowIteratorNext         = nullptr;
    renderRowIteratorFree         = nullptr;
    renderRowIteratorNew          = nullptr;
    renderStateSet                = nullptr;
    renderStateGet                = nullptr;
    renderStateEndUpdate          = nullptr;
    renderStateBeginUpdate        = nullptr;
    renderStateFree               = nullptr;
    renderStateNew                = nullptr;
    formatterFree                 = nullptr;
    formatterFormatBuffer         = nullptr;
    formatterTerminalNew          = nullptr;
    pasteEncode                   = nullptr;
    pasteIsSafe                   = nullptr;
    terminalSelectionFormatBuffer = nullptr;
    terminalScrollViewport        = nullptr;
    terminalSelectAll             = nullptr;
    terminalSelectWord            = nullptr;
    terminalVtWrite               = nullptr;
    trackedGridRefSnapshot        = nullptr;
    trackedGridRefFree            = nullptr;
    terminalGridRefTrack          = nullptr;
    terminalGridRef               = nullptr;
    terminalGet                   = nullptr;
    terminalSet                   = nullptr;
    terminalResize                = nullptr;
    terminalFree                  = nullptr;
    terminalNew                   = nullptr;
    buildInfo                     = nullptr;
    _module.reset();
    _runtimeFile.reset();
}

bool TerminalRuntimeLoader::IsLoaded() const noexcept
{
    return _module && _runtimeFile;
}

const std::wstring& TerminalRuntimeLoader::ErrorText() const noexcept
{
    return _errorText;
}

GhosttyResult TerminalRuntimeLoader::GetTerminalMode(GhosttyTerminal terminal, GhosttyMode mode, bool& value) const noexcept
{
    if (terminalGet == nullptr)
    {
        return GHOSTTY_INVALID_VALUE;
    }

    GhosttyTerminalModeConfig config{};
    config.mode                = mode;
    const GhosttyResult result = terminalGet(terminal, GHOSTTY_TERMINAL_DATA_MODE, &config);
    if (result == GHOSTTY_SUCCESS)
    {
        value = config.value;
    }
    return result;
}

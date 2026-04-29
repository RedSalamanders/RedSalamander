#pragma once

#include "DxUi/DxUi.Internal.h"
#include "DxUi/DxUi.h"
#include "Helpers.h"
#include "WindowMessages.h"

#include <UIAutomation.h>
#include <imm.h>
#include <richedit.h>
#include <wil/com.h>
#include <wil/resource.h>
#include <wincodec.h>
#include <windowsx.h>

#include <algorithm>
#include <array>
#include <cmath>
#include <cstdlib>
#include <filesystem>
#include <format>
#include <iostream>
#include <mutex>
#include <string>
#include <vector>

#pragma comment(lib, "imm32.lib")

inline void Require(bool condition, const char* message)
{
    if (! condition)
    {
        std::cerr << "FAILED: " << message << '\n';
        std::exit(1);
    }
}

inline void RequireColorNear(const D2D1_COLOR_F& actual, const D2D1_COLOR_F& expected, const char* message)
{
    const auto nearlyEqual = [](float a, float b) noexcept { return std::fabs(a - b) <= 0.0001f; };

    if (! nearlyEqual(actual.r, expected.r) || ! nearlyEqual(actual.g, expected.g) || ! nearlyEqual(actual.b, expected.b) ||
        ! nearlyEqual(actual.a, expected.a))
    {
        std::cerr << "FAILED: " << message << '\n';
        std::exit(1);
    }
}

inline void RequireColorDifferent(const D2D1_COLOR_F& lhs, const D2D1_COLOR_F& rhs, const char* message)
{
    const auto nearlyEqual = [](float a, float b) noexcept { return std::fabs(a - b) <= 0.0001f; };

    if (nearlyEqual(lhs.r, rhs.r) && nearlyEqual(lhs.g, rhs.g) && nearlyEqual(lhs.b, rhs.b) && nearlyEqual(lhs.a, rhs.a))
    {
        std::cerr << "FAILED: " << message << '\n';
        std::exit(1);
    }
}

inline void RequirePointNear(const POINT& actual, const POINT& expected, const char* message)
{
    const auto nearlyEqual = [](LONG a, LONG b) noexcept { return std::abs(a - b) <= 1; };

    if (! nearlyEqual(actual.x, expected.x) || ! nearlyEqual(actual.y, expected.y))
    {
        std::cerr << "FAILED: " << message << '\n';
        std::exit(1);
    }
}

[[maybe_unused]] inline void RequireSucceeded(HRESULT hr, const char* message)
{
    if (FAILED(hr))
    {
        std::cerr << "FAILED: " << message << " hr=0x" << std::hex << static_cast<unsigned long>(hr) << std::dec << '\n';
        std::exit(1);
    }
}

inline void RequireRectNear(const RECT& actual, const RECT& expected, const char* message)
{
    const auto nearlyEqual = [](LONG a, LONG b) noexcept { return std::abs(a - b) <= 1; };

    if (! nearlyEqual(actual.left, expected.left) || ! nearlyEqual(actual.top, expected.top) || ! nearlyEqual(actual.right, expected.right) ||
        ! nearlyEqual(actual.bottom, expected.bottom))
    {
        std::cerr << "FAILED: " << message << '\n';
        std::exit(1);
    }
}

inline void RequireRectHasArea(const D2D1_RECT_F& rect, const char* message)
{
    if (rect.right <= rect.left || rect.bottom <= rect.top)
    {
        std::cerr << "FAILED: " << message << '\n';
        std::exit(1);
    }
}

inline void RequireFloatNear(float actual, float expected, float epsilon, const char* message)
{
    if (std::fabs(actual - expected) > epsilon)
    {
        std::cerr << "FAILED: " << message << '\n';
        std::exit(1);
    }
}

inline bool& DxUiWriteBaselinesFlag() noexcept
{
    static bool value = false;
    return value;
}

inline void SetDxUiWriteBaselines(bool value) noexcept
{
    DxUiWriteBaselinesFlag() = value;
}

inline bool ShouldWriteDxUiBaselines() noexcept
{
    return DxUiWriteBaselinesFlag();
}

inline std::filesystem::path FindRepoRootForDxUiTests()
{
    std::filesystem::path current = std::filesystem::current_path();
    for (size_t depth = 0u; depth < 8u; ++depth)
    {
        if (std::filesystem::exists(current / L"RedSalamander.sln") && std::filesystem::exists(current / L"Tests" / L"DxUiTests"))
        {
            return current;
        }

        if (! current.has_parent_path())
        {
            break;
        }
        current = current.parent_path();
    }

    std::cerr << "FAILED: unable to locate repo root for DxUi baseline tests\n";
    std::exit(1);
}

inline std::filesystem::path GetDxUiBaselineDirectory()
{
    return FindRepoRootForDxUiTests() / L"Tests" / L"DxUiTests" / L"Baselines";
}

inline std::filesystem::path GetDxUiBaselinePath(std::wstring_view fileName)
{
    return GetDxUiBaselineDirectory() / std::filesystem::path(fileName);
}

inline wil::com_ptr<IWICImagingFactory> CreateWicFactoryForTest()
{
    static thread_local const HRESULT hrCoInit = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    Require(hrCoInit == S_OK || hrCoInit == S_FALSE || hrCoInit == RPC_E_CHANGED_MODE, "WIC factory COM initialization succeeds");

    wil::com_ptr<IWICImagingFactory> factory;
    HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory2, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(factory.addressof()));
    if (FAILED(hr) || ! factory)
    {
        hr = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(factory.addressof()));
    }
    RequireSucceeded(hr, "WIC imaging factory created for baseline tests");
    Require(factory != nullptr, "WIC imaging factory instance created for baseline tests");
    return factory;
}

inline bool SaveWindowHostBitmapCaptureAsPngForTest(const std::filesystem::path& path, const RedSalamander::DxUi::WindowHostBitmapCapture& capture)
{
    if (capture.widthPx == 0u || capture.heightPx == 0u || capture.bgraPixels.empty())
    {
        return false;
    }

    std::error_code ec;
    std::filesystem::create_directories(path.parent_path(), ec);
    if (ec)
    {
        return false;
    }

    const auto factory = CreateWicFactoryForTest();

    wil::com_ptr<IWICStream> stream;
    if (FAILED(factory->CreateStream(stream.addressof())) || ! stream)
    {
        return false;
    }
    if (FAILED(stream->InitializeFromFilename(path.c_str(), GENERIC_WRITE)))
    {
        return false;
    }

    wil::com_ptr<IWICBitmapEncoder> encoder;
    if (FAILED(factory->CreateEncoder(GUID_ContainerFormatPng, nullptr, encoder.addressof())) || ! encoder)
    {
        return false;
    }
    if (FAILED(encoder->Initialize(stream.get(), WICBitmapEncoderNoCache)))
    {
        return false;
    }

    wil::com_ptr<IWICBitmapFrameEncode> frame;
    wil::com_ptr<IPropertyBag2> properties;
    if (FAILED(encoder->CreateNewFrame(frame.addressof(), properties.addressof())) || ! frame)
    {
        return false;
    }
    if (FAILED(frame->Initialize(properties.get())) || FAILED(frame->SetSize(capture.widthPx, capture.heightPx)))
    {
        return false;
    }

    WICPixelFormatGUID pixelFormat = GUID_WICPixelFormat32bppBGRA;
    if (FAILED(frame->SetPixelFormat(&pixelFormat)))
    {
        return false;
    }
    if (pixelFormat != GUID_WICPixelFormat32bppBGRA)
    {
        return false;
    }

    const UINT stride = capture.widthPx * 4u;
    if (FAILED(frame->WritePixels(capture.heightPx,
                                  stride,
                                  static_cast<UINT>(capture.bgraPixels.size()),
                                  const_cast<BYTE*>(reinterpret_cast<const BYTE*>(capture.bgraPixels.data())))))
    {
        return false;
    }

    return SUCCEEDED(frame->Commit()) && SUCCEEDED(encoder->Commit());
}

inline bool LoadWindowHostBitmapCaptureFromPngForTest(const std::filesystem::path& path, RedSalamander::DxUi::WindowHostBitmapCapture& capture)
{
    capture = {};
    if (! std::filesystem::exists(path))
    {
        return false;
    }

    const auto factory = CreateWicFactoryForTest();

    wil::com_ptr<IWICBitmapDecoder> decoder;
    if (FAILED(factory->CreateDecoderFromFilename(path.c_str(), nullptr, GENERIC_READ, WICDecodeMetadataCacheOnLoad, decoder.addressof())) || ! decoder)
    {
        return false;
    }

    wil::com_ptr<IWICBitmapFrameDecode> frame;
    if (FAILED(decoder->GetFrame(0u, frame.addressof())) || ! frame)
    {
        return false;
    }

    UINT width  = 0u;
    UINT height = 0u;
    if (FAILED(frame->GetSize(&width, &height)) || width == 0u || height == 0u)
    {
        return false;
    }

    wil::com_ptr<IWICFormatConverter> converter;
    if (FAILED(factory->CreateFormatConverter(converter.addressof())) || ! converter)
    {
        return false;
    }
    if (FAILED(converter->Initialize(frame.get(), GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherTypeNone, nullptr, 0.0f, WICBitmapPaletteTypeCustom)))
    {
        return false;
    }

    capture.widthPx  = width;
    capture.heightPx = height;
    capture.bgraPixels.resize(static_cast<size_t>(width) * static_cast<size_t>(height) * 4u);
    const UINT stride = width * 4u;
    return SUCCEEDED(converter->CopyPixels(nullptr, stride, static_cast<UINT>(capture.bgraPixels.size()), reinterpret_cast<BYTE*>(capture.bgraPixels.data())));
}

struct BitmapComparisonStats
{
    size_t differingPixels  = 0u;
    size_t totalPixels      = 0u;
    uint8_t maxChannelDelta = 0u;

    [[nodiscard]] double DifferenceRatio() const noexcept
    {
        return totalPixels == 0u ? 1.0 : static_cast<double>(differingPixels) / static_cast<double>(totalPixels);
    }
};

inline BitmapComparisonStats CompareWindowHostBitmapCapturesForTest(const RedSalamander::DxUi::WindowHostBitmapCapture& actual,
                                                                    const RedSalamander::DxUi::WindowHostBitmapCapture& expected,
                                                                    uint8_t perChannelTolerance = 8u) noexcept
{
    BitmapComparisonStats stats{};
    if (actual.widthPx != expected.widthPx || actual.heightPx != expected.heightPx)
    {
        return stats;
    }

    stats.totalPixels = static_cast<size_t>(actual.widthPx) * static_cast<size_t>(actual.heightPx);
    for (size_t pixelIndex = 0u; pixelIndex < stats.totalPixels; ++pixelIndex)
    {
        const size_t base     = pixelIndex * 4u;
        uint8_t pixelMaxDelta = 0u;
        for (size_t channelIndex = 0u; channelIndex < 4u; ++channelIndex)
        {
            const uint8_t actualValue   = actual.bgraPixels[base + channelIndex];
            const uint8_t expectedValue = expected.bgraPixels[base + channelIndex];
            const uint8_t delta =
                actualValue > expectedValue ? static_cast<uint8_t>(actualValue - expectedValue) : static_cast<uint8_t>(expectedValue - actualValue);
            pixelMaxDelta = (std::max)(pixelMaxDelta, delta);
        }

        stats.maxChannelDelta = (std::max)(stats.maxChannelDelta, pixelMaxDelta);
        if (pixelMaxDelta > perChannelTolerance)
        {
            ++stats.differingPixels;
        }
    }

    return stats;
}

inline void VerifyOrUpdateBaselineForTest(const char* context,
                                          std::wstring_view fileName,
                                          const RedSalamander::DxUi::WindowHostBitmapCapture& actual,
                                          double maxDifferenceRatio   = 0.02,
                                          uint8_t perChannelTolerance = 8u)
{
    const std::filesystem::path baselinePath = GetDxUiBaselinePath(fileName);
    if (ShouldWriteDxUiBaselines())
    {
        Require(SaveWindowHostBitmapCaptureAsPngForTest(baselinePath, actual), context);
        return;
    }

    RedSalamander::DxUi::WindowHostBitmapCapture expected;
    Require(LoadWindowHostBitmapCaptureFromPngForTest(baselinePath, expected), context);
    Require(expected.widthPx == actual.widthPx && expected.heightPx == actual.heightPx, context);

    const BitmapComparisonStats stats = CompareWindowHostBitmapCapturesForTest(actual, expected, perChannelTolerance);
    if (stats.DifferenceRatio() > maxDifferenceRatio)
    {
        const std::filesystem::path actualPath = baselinePath.parent_path() / L"_Actual" / baselinePath.filename();
        static_cast<void>(SaveWindowHostBitmapCaptureAsPngForTest(actualPath, actual));
        std::cerr << "FAILED: " << context << " diffRatio=" << stats.DifferenceRatio() << " maxChannelDelta=" << static_cast<int>(stats.maxChannelDelta)
                  << " actual=" << actualPath.string() << '\n';
        std::exit(1);
    }
}

inline size_t MapRichEditBridgeIndexToLfIndexForTest(std::wstring_view text, size_t bridgeIndex)
{
    size_t lfIndex       = 0u;
    size_t richEditIndex = 0u;
    while (lfIndex < text.size())
    {
        if (text[lfIndex] == L'\n')
        {
            if (bridgeIndex <= richEditIndex)
            {
                return lfIndex;
            }
            ++richEditIndex; // after CR, before LF still maps to the logical pre-newline position
            if (bridgeIndex <= richEditIndex)
            {
                return lfIndex;
            }
            ++richEditIndex;
            ++lfIndex;
        }
        else
        {
            if (bridgeIndex <= richEditIndex)
            {
                return lfIndex;
            }
            ++richEditIndex;
            ++lfIndex;
        }
    }
    return lfIndex;
}

inline bool BridgeCollapsedCaretMatchesVisibleIndexForTest(std::wstring_view text, size_t bridgeIndex, size_t visibleIndex)
{
    return bridgeIndex == visibleIndex || MapRichEditBridgeIndexToLfIndexForTest(text, bridgeIndex) == visibleIndex;
}

inline bool BridgeCollapsedCaretMatchesVisibleTrailingNewlineBoundaryForTest(std::wstring_view text, size_t bridgeIndex, size_t visibleIndex)
{
    const size_t mappedIndex = MapRichEditBridgeIndexToLfIndexForTest(text, bridgeIndex);
    return bridgeIndex == visibleIndex || mappedIndex == visibleIndex || mappedIndex + 1u == visibleIndex;
}

// Clipboard listeners and sync providers can momentarily race the first WM_PASTE
// observation on desktop test runs, so clipboard-driven bridge tests retry the
// full interaction a few times instead of failing on the first transient miss.
template <typename TAction> inline bool RetryClipboardSensitiveBridgeAction(TAction action)
{
    for (int attempt = 0; attempt < 4; ++attempt)
    {
        if (action())
        {
            return true;
        }

        Sleep(20);
    }

    return false;
}

inline std::optional<std::wstring> ReadClipboardUnicodeTextForTest(HWND ownerWindow);

inline bool DispatchQueuedMessageForTest(const MSG& msg)
{
    if (msg.hwnd != nullptr && IsWindow(msg.hwnd) == FALSE)
    {
        std::cerr << "    [TRACE] dropping queued message for destroyed hwnd: msg=" << msg.message << " hwnd=" << msg.hwnd << " wp=" << msg.wParam
                  << " lp=" << msg.lParam << '\n'
                  << std::flush;
        return false;
    }

    TranslateMessage(&msg);
    DispatchMessageW(&msg);
    return true;
}

inline wil::unique_hwnd CreateClipboardOwnerWindowForTest()
{
    HWND hwnd =
        CreateWindowExW(0, L"STATIC", L"DxUiTestsClipboardOwner", WS_OVERLAPPED, -32000, -32000, 16, 16, nullptr, nullptr, GetModuleHandleW(nullptr), nullptr);
    Require(hwnd != nullptr, "clipboard owner window created");
    return wil::unique_hwnd(hwnd);
}

inline bool SetClipboardUnicodeTextForTest(HWND ownerWindow, std::wstring_view text)
{
    const auto openClipboardWithRetries = [ownerWindow]() noexcept
    {
        if (! ownerWindow)
        {
            return false;
        }

        for (int attempt = 0; attempt < 20; ++attempt)
        {
            if (OpenClipboard(ownerWindow) != 0)
            {
                return true;
            }
            if (GetOpenClipboardWindow() == nullptr)
            {
                static_cast<void>(CloseClipboard());
            }
            Sleep(10);
        }
        return false;
    };

    if (! openClipboardWithRetries())
    {
        return RedSalamander::DxUi::DebugSetClipboardFallbackText(text);
    }
    auto closeClipboard = wil::scope_exit([&] { CloseClipboard(); });

    if (EmptyClipboard() == 0)
    {
        return RedSalamander::DxUi::DebugSetClipboardFallbackText(text);
    }

    const size_t bytes = (text.size() + 1u) * sizeof(wchar_t);
    wil::unique_hglobal memory(GlobalAlloc(GMEM_MOVEABLE, bytes));
    if (! memory)
    {
        return RedSalamander::DxUi::DebugSetClipboardFallbackText(text);
    }

    HGLOBAL memoryHandle = memory.get();
    void* lockedMemory   = GlobalLock(memoryHandle);
    if (! lockedMemory)
    {
        return RedSalamander::DxUi::DebugSetClipboardFallbackText(text);
    }
    auto* out = static_cast<wchar_t*>(lockedMemory);
    std::copy(text.begin(), text.end(), out);
    out[text.size()] = L'\0';

    if (GlobalUnlock(memoryHandle) == FALSE)
    {
        const DWORD unlockError = GetLastError();
        if (unlockError != NO_ERROR)
        {
            return RedSalamander::DxUi::DebugSetClipboardFallbackText(text);
        }
    }
    if (SetClipboardData(CF_UNICODETEXT, memoryHandle) == nullptr)
    {
        return RedSalamander::DxUi::DebugSetClipboardFallbackText(text);
    }

    memory.release();
    closeClipboard.release();
    static_cast<void>(CloseClipboard());

    for (int attempt = 0; attempt < 20; ++attempt)
    {
        const std::optional<std::wstring> clipboardText = ReadClipboardUnicodeTextForTest(ownerWindow);
        if (clipboardText && clipboardText.value() == text)
        {
            static_cast<void>(RedSalamander::DxUi::DebugSetClipboardFallbackText(text));
            return true;
        }

        Sleep(10);
    }

    return RedSalamander::DxUi::DebugSetClipboardFallbackText(text);
}

inline std::optional<std::wstring> ReadClipboardUnicodeTextForTest(HWND ownerWindow)
{
    const auto openClipboardWithRetries = [ownerWindow]() noexcept
    {
        if (! ownerWindow)
        {
            return false;
        }

        for (int attempt = 0; attempt < 20; ++attempt)
        {
            if (OpenClipboard(ownerWindow) != 0)
            {
                return true;
            }
            if (GetOpenClipboardWindow() == nullptr)
            {
                static_cast<void>(CloseClipboard());
            }
            Sleep(10);
        }
        return false;
    };

    if (! openClipboardWithRetries())
    {
        return RedSalamander::DxUi::DebugReadClipboardFallbackText();
    }
    const auto closeClipboard = wil::scope_exit([&] { CloseClipboard(); });

    HANDLE handle = GetClipboardData(CF_UNICODETEXT);
    if (! handle)
    {
        return RedSalamander::DxUi::DebugReadClipboardFallbackText();
    }

    LPCWSTR text = static_cast<LPCWSTR>(GlobalLock(handle));
    if (! text)
    {
        return RedSalamander::DxUi::DebugReadClipboardFallbackText();
    }
    const auto unlockMemory = wil::scope_exit([&] { GlobalUnlock(handle); });
    std::wstring clipboardText(text);
    static_cast<void>(RedSalamander::DxUi::DebugSetClipboardFallbackText(clipboardText));
    return clipboardText;
}

inline D2D1_COLOR_F BlendForTest(const D2D1_COLOR_F& a, const D2D1_COLOR_F& b, float t) noexcept
{
    const float clamped = std::clamp(t, 0.0f, 1.0f);
    return D2D1::ColorF(a.r + ((b.r - a.r) * clamped), a.g + ((b.g - a.g) * clamped), a.b + ((b.b - a.b) * clamped), a.a + ((b.a - a.a) * clamped));
}

inline D2D1_COLOR_F ChooseContrastingTextColorForTest(const D2D1_COLOR_F& background) noexcept
{
    const float luminance = background.r * 0.2126f + background.g * 0.7152f + background.b * 0.0722f;
    return luminance >= 0.55f ? D2D1::ColorF(0.06f, 0.06f, 0.06f, 1.0f) : D2D1::ColorF(0.98f, 0.98f, 0.98f, 1.0f);
}

[[maybe_unused]] inline uint32_t PackColorForTest(const D2D1_COLOR_F& color) noexcept
{
    const auto packChannel = [](float component) noexcept { return static_cast<uint32_t>(std::lround(std::clamp(component, 0.0f, 1.0f) * 255.0f)); };

    return (packChannel(color.a) << 24u) | (packChannel(color.r) << 16u) | (packChannel(color.g) << 8u) | packChannel(color.b);
}

#if defined(ENABLE_TESTS)
[[nodiscard]] inline std::wstring ReadProviderStringProperty(IRawElementProviderSimple& provider, PROPERTYID propertyId, const char* context)
{
    VARIANT value{};
    VariantInit(&value);
    const auto clearValue = wil::scope_exit([&] { VariantClear(&value); });
    RequireSucceeded(provider.GetPropertyValue(propertyId, &value), context);
    Require(value.vt == VT_BSTR, context);
    return value.bstrVal ? std::wstring(value.bstrVal, SysStringLen(value.bstrVal)) : std::wstring{};
}

[[nodiscard]] inline LONG ReadProviderLongProperty(IRawElementProviderSimple& provider, PROPERTYID propertyId, const char* context)
{
    VARIANT value{};
    VariantInit(&value);
    const auto clearValue = wil::scope_exit([&] { VariantClear(&value); });
    RequireSucceeded(provider.GetPropertyValue(propertyId, &value), context);
    Require(value.vt == VT_I4, context);
    return value.lVal;
}

[[nodiscard]] inline bool ReadProviderBoolProperty(IRawElementProviderSimple& provider, PROPERTYID propertyId, const char* context)
{
    VARIANT value{};
    VariantInit(&value);
    const auto clearValue = wil::scope_exit([&] { VariantClear(&value); });
    RequireSucceeded(provider.GetPropertyValue(propertyId, &value), context);
    Require(value.vt == VT_BOOL, context);
    return value.boolVal == VARIANT_TRUE;
}

[[nodiscard]] inline std::vector<std::wstring> ReadSelectionProviderNames(ISelectionProvider& selectionProvider, const char* context)
{
    SAFEARRAY* selectionArray = nullptr;
    RequireSucceeded(selectionProvider.GetSelection(&selectionArray), context);
    Require(selectionArray != nullptr, context);
    const auto destroyArray = wil::scope_exit([&] { SafeArrayDestroy(selectionArray); });

    LONG lowerBound = 0;
    LONG upperBound = -1;
    RequireSucceeded(SafeArrayGetLBound(selectionArray, 1, &lowerBound), context);
    RequireSucceeded(SafeArrayGetUBound(selectionArray, 1, &upperBound), context);

    std::vector<std::wstring> names;
    if (upperBound < lowerBound)
    {
        return names;
    }

    names.reserve(static_cast<size_t>(upperBound - lowerBound + 1));
    for (LONG index = lowerBound; index <= upperBound; ++index)
    {
        wil::com_ptr_nothrow<IUnknown> unknown;
        RequireSucceeded(SafeArrayGetElement(selectionArray, &index, unknown.put_void()), context);
        wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
        RequireSucceeded(unknown.query_to(simple.put()), context);
        names.push_back(ReadProviderStringProperty(*simple.get(), UIA_NamePropertyId, context));
    }

    return names;
}

[[nodiscard]] inline std::vector<std::wstring> ReadProviderArrayNames(SAFEARRAY* providerArray, const char* context)
{
    Require(providerArray != nullptr, context);

    LONG lowerBound = 0;
    LONG upperBound = -1;
    RequireSucceeded(SafeArrayGetLBound(providerArray, 1, &lowerBound), context);
    RequireSucceeded(SafeArrayGetUBound(providerArray, 1, &upperBound), context);

    std::vector<std::wstring> names;
    if (upperBound < lowerBound)
    {
        return names;
    }

    names.reserve(static_cast<size_t>(upperBound - lowerBound + 1));
    for (LONG index = lowerBound; index <= upperBound; ++index)
    {
        wil::com_ptr_nothrow<IUnknown> unknown;
        RequireSucceeded(SafeArrayGetElement(providerArray, &index, unknown.put_void()), context);
        wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
        RequireSucceeded(unknown.query_to(simple.put()), context);
        names.push_back(ReadProviderStringProperty(*simple.get(), UIA_NamePropertyId, context));
    }

    return names;
}

[[nodiscard]] inline wil::com_ptr_nothrow<IRawElementProviderFragment> GetProviderAtDipPoint(
    HWND hostHwnd, RedSalamander::DxUi::WindowHost& host, IRawElementProviderFragmentRoot& rootProvider, float xDip, float yDip, const char* context)
{
    POINT pointPx{
        static_cast<LONG>(std::lround(host.DipsToPixels(xDip))),
        static_cast<LONG>(std::lround(host.DipsToPixels(yDip))),
    };
    Require(ClientToScreen(hostHwnd, &pointPx) != FALSE, context);

    wil::com_ptr_nothrow<IRawElementProviderFragment> provider;
    RequireSucceeded(rootProvider.ElementProviderFromPoint(static_cast<double>(pointPx.x), static_cast<double>(pointPx.y), provider.put()), context);
    Require(provider != nullptr, context);
    return provider;
}
#endif

class AttachedHostWindow final
{
public:
    AttachedHostWindow(const AttachedHostWindow&)            = delete;
    AttachedHostWindow& operator=(const AttachedHostWindow&) = delete;
    AttachedHostWindow(AttachedHostWindow&&)                 = delete;
    AttachedHostWindow& operator=(AttachedHostWindow&&)      = delete;

    AttachedHostWindow()
    {
        static_cast<void>(EnsureWindowClass());
        HWND hwnd =
            CreateWindowExW(0, kWindowClassName, L"DxUiTestsHost", WS_OVERLAPPED, -32000, -32000, 320, 200, nullptr, nullptr, GetModuleHandleW(nullptr), this);
        Require(hwnd != nullptr, "attached host window created");
        _hwnd.reset(hwnd);
        Require(_host.Attach(_hwnd.get()), "attached host window host attached");
    }

    ~AttachedHostWindow()
    {
        std::cerr << "    [TRACE] attached host dtor: begin\n" << std::flush;
        _host.Detach();
        std::cerr << "    [TRACE] attached host dtor: host detached\n" << std::flush;
        _hwnd.reset();
        std::cerr << "    [TRACE] attached host dtor: hwnd reset\n" << std::flush;
    }

    [[nodiscard]] RedSalamander::DxUi::WindowHost& Host() noexcept
    {
        return _host;
    }

    [[nodiscard]] HWND Hwnd() const noexcept
    {
        return _hwnd.get();
    }

    void PumpMessages() const
    {
        MSG msg{};
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) != FALSE)
        {
            static_cast<void>(DispatchQueuedMessageForTest(msg));
        }
    }

private:
    static constexpr PCWSTR kWindowClassName = L"RedSalamander.DxUiTests.AttachedHostWindow";

    static ATOM EnsureWindowClass()
    {
        static const ATOM atom = []() noexcept
        {
            WNDCLASSW windowClass{};
            windowClass.lpfnWndProc   = &AttachedHostWindow::WndProc;
            windowClass.hInstance     = GetModuleHandleW(nullptr);
            windowClass.lpszClassName = kWindowClassName;
            return RegisterClassW(&windowClass);
        }();
        return atom;
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
    {
        if (msg == WM_NCCREATE)
        {
            const auto* createStruct = reinterpret_cast<const CREATESTRUCTW*>(lp);
            auto* self               = static_cast<AttachedHostWindow*>(createStruct->lpCreateParams);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            return TRUE;
        }

        auto* self = reinterpret_cast<AttachedHostWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (self)
        {
            bool handled         = false;
            const LRESULT result = self->_host.HandleMessage(hwnd, msg, wp, lp, handled);
            if (msg == WM_NCDESTROY)
            {
                self->_host.Detach();
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            }
            if (handled)
            {
                return result;
            }
        }

        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    wil::unique_hwnd _hwnd;
    RedSalamander::DxUi::WindowHost _host;
};

class ClipboardHostWindow final
{
public:
    ClipboardHostWindow(const ClipboardHostWindow&)            = delete;
    ClipboardHostWindow& operator=(const ClipboardHostWindow&) = delete;
    ClipboardHostWindow(ClipboardHostWindow&&)                 = delete;
    ClipboardHostWindow& operator=(ClipboardHostWindow&&)      = delete;

    ClipboardHostWindow() : _hwnd(CreateClipboardOwnerWindowForTest())
    {
        Require(_host.Attach(_hwnd.get()), "clipboard host attached");
    }

    ~ClipboardHostWindow()
    {
        _host.Detach();
    }

    [[nodiscard]] RedSalamander::DxUi::WindowHost& Host() noexcept
    {
        return _host;
    }

    [[nodiscard]] HWND Hwnd() const noexcept
    {
        return _hwnd.get();
    }

private:
    wil::unique_hwnd _hwnd;
    RedSalamander::DxUi::WindowHost _host;
};

class ExposedTextField final : public RedSalamander::DxUi::TextField
{
public:
    using RedSalamander::DxUi::TextField::ExportTextInputBridgeState;
    using RedSalamander::DxUi::TextField::ImportTextInputBridgeState;
    using RedSalamander::DxUi::TextField::TextField;
};

class ExposedButton final : public RedSalamander::DxUi::Button
{
public:
    using RedSalamander::DxUi::Button::Button;
    using RedSalamander::DxUi::Button::IsPressed;
    using RedSalamander::DxUi::Button::OnCaptureLost;
    using RedSalamander::DxUi::Button::OnFocusChanged;
    using RedSalamander::DxUi::Button::OnHoverChanged;
    using RedSalamander::DxUi::Button::SetPressed;
    using RedSalamander::DxUi::Button::Tick;
};

struct RecordingContextMenuInvocation
{
    size_t count = 0u;
    POINT lastPoint{};
    bool lastKeyboardInvocation = false;

    void Record(POINT screenPoint, bool keyboardInvocation) noexcept
    {
        ++count;
        lastPoint              = screenPoint;
        lastKeyboardInvocation = keyboardInvocation;
    }
};

[[nodiscard]] inline HWND FindTextInputBridgeEdit(HWND hostHwnd)
{
    if (HWND bridge = FindWindowExW(hostHwnd, nullptr, L"DxUiTextInputBridgeWindow", nullptr))
    {
        return bridge;
    }
    if (HWND richEdit = FindWindowExW(hostHwnd, nullptr, MSFTEDIT_CLASS, nullptr))
    {
        return richEdit;
    }
    return FindWindowExW(hostHwnd, nullptr, L"EDIT", nullptr);
}

[[nodiscard]] inline RECT ComputeExpectedBridgeRect(HWND hostHwnd, RedSalamander::DxUi::WindowHost& host, const D2D1_RECT_F& boundsDip)
{
    constexpr int kExpectedBridgeMinWidthPx  = 64;
    constexpr int kExpectedBridgeMinHeightPx = 32;

    POINT clientOrigin{};
    Require(ClientToScreen(hostHwnd, &clientOrigin) != FALSE, "host client origin is available for text bridge rect");

    const LONG left   = clientOrigin.x + static_cast<LONG>(std::lround(host.DipsToPixels(boundsDip.left)));
    const LONG top    = clientOrigin.y + static_cast<LONG>(std::lround(host.DipsToPixels(boundsDip.top)));
    const LONG width  = std::max<LONG>(kExpectedBridgeMinWidthPx, static_cast<LONG>(std::lround(host.DipsToPixels(boundsDip.right - boundsDip.left))));
    const LONG height = std::max<LONG>(kExpectedBridgeMinHeightPx, static_cast<LONG>(std::lround(host.DipsToPixels(boundsDip.bottom - boundsDip.top))));
    const LONG right  = left + width;
    const LONG bottom = top + height;
    return RECT{left, top, right, bottom};
}

[[nodiscard]] inline std::wstring ReadBridgeTextContent(HWND bridgeEdit)
{
    const int bridgeLength = GetWindowTextLengthW(bridgeEdit);
    Require(bridgeLength >= 0, "can read the hidden bridge text length");
    std::wstring bridgeText(static_cast<size_t>(bridgeLength) + 1u, L'\0');
    const int copied = GetWindowTextW(bridgeEdit, bridgeText.data(), static_cast<int>(bridgeText.size()));
    bridgeText.resize(static_cast<size_t>((std::max)(0, copied)));
    return bridgeText;
}

constexpr std::wstring_view kLogicalNewlineClipboardTextForTest           = L"alpha\nbeta";
constexpr size_t kLogicalNewlineClipboardSelectionStartForTest            = 2u;
constexpr size_t kLogicalNewlineClipboardSelectionEndForTest              = 7u;
constexpr std::wstring_view kLogicalNewlineClipboardSelectedTextForTest   = L"pha\nb";
constexpr std::wstring_view kLogicalNewlineClipboardCutResultForTest      = L"aleta";
constexpr std::wstring_view kLogicalNewlinePasteClipboardTextForTest      = L"\r\nZ";
constexpr std::wstring_view kLogicalNewlinePasteInsertedTextForTest       = L"\nZ";
constexpr std::wstring_view kLogicalNewlinePasteResultForTest             = L"al\nZeta";
constexpr std::wstring_view kWrappedMultilineClipboardTextForTest         = L"alpha bravo charlie delta echo foxtrot golf hotel";
constexpr size_t kWrappedMultilineClipboardSelectionStartForTest          = 6u;
constexpr size_t kWrappedMultilineClipboardSelectionEndForTest            = 19u;
constexpr std::wstring_view kWrappedMultilineClipboardSelectedTextForTest = L"bravo charlie";
constexpr std::wstring_view kWrappedMultilinePasteClipboardTextForTest    = L"XYZ";
constexpr std::wstring_view kWrappedMultilinePasteResultForTest           = L"alpha XYZ delta echo foxtrot golf hotel";

[[nodiscard]] inline std::wstring RemoveSelectionForTest(std::wstring_view text, size_t selectionStart, size_t selectionEnd)
{
    std::wstring result(text.substr(0u, selectionStart));
    result.append(text.substr(selectionEnd));
    return result;
}

inline void ImportLogicalNewlineClipboardSelectionForTest(RedSalamander::DxUi::WindowHost& host, ExposedTextField& field, const char* message)
{
    RedSalamander::DxUi::TextInputBridgeState state;
    state.text                 = field.GetText();
    state.caretIndex           = kLogicalNewlineClipboardSelectionEndForTest;
    state.selectionAnchorIndex = kLogicalNewlineClipboardSelectionStartForTest;
    state.firstVisibleLine     = 0u;
    state.multiline            = true;
    Require(field.ImportTextInputBridgeState(host, state, false), message);
}

inline void RequireLogicalNewlineClipboardVisibleSelectionForTest(const RedSalamander::DxUi::TextInputBridgeState& state, const char* context)
{
    Require(state.selectionAnchorIndex.has_value(), context);
    const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(visibleSelectionStart == kLogicalNewlineClipboardSelectionStartForTest && visibleSelectionEnd == kLogicalNewlineClipboardSelectionEndForTest,
            context);
}

inline void RequireLogicalNewlineClipboardBridgeSelectionForTest(HWND bridgeEdit, const char* message)
{
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    const size_t mappedSelectionStart = MapRichEditBridgeIndexToLfIndexForTest(kLogicalNewlineClipboardTextForTest, static_cast<size_t>(selectionStart));
    const size_t mappedSelectionEnd   = MapRichEditBridgeIndexToLfIndexForTest(kLogicalNewlineClipboardTextForTest, static_cast<size_t>(selectionEnd));
    const bool selectionEndIsBridgeAligned =
        mappedSelectionEnd == kLogicalNewlineClipboardSelectionEndForTest || mappedSelectionEnd + 1u == kLogicalNewlineClipboardSelectionEndForTest;
    Require(mappedSelectionStart == kLogicalNewlineClipboardSelectionStartForTest && selectionEndIsBridgeAligned, message);
}

inline void ImportWrappedMultilineClipboardSelectionForTest(RedSalamander::DxUi::WindowHost& host, ExposedTextField& field, const char* message)
{
    RedSalamander::DxUi::TextInputBridgeState state;
    state.text                 = field.GetText();
    state.caretIndex           = kWrappedMultilineClipboardSelectionEndForTest;
    state.selectionAnchorIndex = kWrappedMultilineClipboardSelectionStartForTest;
    state.firstVisibleLine     = 0u;
    state.multiline            = true;
    Require(field.ImportTextInputBridgeState(host, state, false), message);
}

inline void RequireWrappedMultilineClipboardVisibleSelectionForTest(const RedSalamander::DxUi::TextInputBridgeState& state, const char* context)
{
    Require(state.selectionAnchorIndex.has_value(), context);
    const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(visibleSelectionStart == kWrappedMultilineClipboardSelectionStartForTest && visibleSelectionEnd == kWrappedMultilineClipboardSelectionEndForTest,
            context);
}

inline void RequireWrappedMultilineClipboardBridgeSelectionForTest(HWND bridgeEdit, const char* message)
{
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == kWrappedMultilineClipboardSelectionStartForTest &&
                static_cast<size_t>(selectionEnd) == kWrappedMultilineClipboardSelectionEndForTest,
            message);
}

[[nodiscard]] inline POINT ClientPointToScreenForTest(HWND hwnd, POINT point, const char* context)
{
    Require(ClientToScreen(hwnd, &point) != FALSE, context);
    return point;
}

[[nodiscard]] inline RECT GetTextBridgeCaretClientRectForTest(HWND bridgeEdit)
{
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));

    std::array<wchar_t, 16> className{};
    const int classLength = GetClassNameW(bridgeEdit, className.data(), static_cast<int>(className.size()));
    const bool richEdit   = classLength > 0 && _wcsicmp(className.data(), MSFTEDIT_CLASS) == 0;

    std::optional<POINT> caretPoint;
    if (richEdit)
    {
        POINTL richEditPoint{};
        if (SendMessageW(bridgeEdit, EM_POSFROMCHAR, reinterpret_cast<WPARAM>(&richEditPoint), static_cast<LPARAM>(selectionEnd)) != -1)
        {
            caretPoint = POINT{richEditPoint.x, richEditPoint.y};
        }
    }
    else
    {
        const LRESULT position = SendMessageW(bridgeEdit, EM_POSFROMCHAR, static_cast<WPARAM>(selectionEnd), 0);
        caretPoint             = POINT{GET_X_LPARAM(position), GET_Y_LPARAM(position)};
    }

    if (! caretPoint.has_value())
    {
        GUITHREADINFO guiThreadInfo{};
        guiThreadInfo.cbSize = sizeof(guiThreadInfo);
        Require(GetGUIThreadInfo(0, &guiThreadInfo) != FALSE && guiThreadInfo.hwndCaret != nullptr,
                "text bridge test can read either EM_POSFROMCHAR or GUI-thread caret geometry");

        RECT caretRect = guiThreadInfo.rcCaret;
        if (guiThreadInfo.hwndCaret != bridgeEdit)
        {
            MapWindowPoints(guiThreadInfo.hwndCaret, bridgeEdit, reinterpret_cast<POINT*>(&caretRect), 2);
        }

        if (caretRect.right <= caretRect.left)
        {
            caretRect.right = caretRect.left + 1;
        }
        if (caretRect.bottom <= caretRect.top)
        {
            caretRect.bottom = caretRect.top + std::max(12L, static_cast<LONG>(GetSystemMetrics(SM_CYCURSOR)));
        }

        return caretRect;
    }

    RECT textRect{};
    if (SendMessageW(bridgeEdit, EM_GETRECT, 0, reinterpret_cast<LPARAM>(&textRect)) == 0)
    {
        GetClientRect(bridgeEdit, &textRect);
    }

    const LONG caretHeight = std::max<LONG>(12, std::max(1L, textRect.bottom - textRect.top));
    return RECT{caretPoint.value().x, caretPoint.value().y, caretPoint.value().x + 1, caretPoint.value().y + caretHeight};
}

[[nodiscard]] inline std::optional<COMPOSITIONFORM> ReadTextBridgeCompositionFormForTest(HWND bridgeEdit)
{
    if (! bridgeEdit)
    {
        return std::nullopt;
    }

    HIMC inputContext = ImmGetContext(bridgeEdit);
    if (! inputContext)
    {
        return std::nullopt;
    }
    const auto releaseContext = wil::scope_exit([&] { ImmReleaseContext(bridgeEdit, inputContext); });

    COMPOSITIONFORM compositionForm{};
    if (ImmGetCompositionWindow(inputContext, &compositionForm) == FALSE)
    {
        return std::nullopt;
    }

    return compositionForm;
}

[[nodiscard]] inline std::optional<CANDIDATEFORM> ReadTextBridgeCandidateFormForTest(HWND bridgeEdit, DWORD index)
{
    if (! bridgeEdit)
    {
        return std::nullopt;
    }

    HIMC inputContext = ImmGetContext(bridgeEdit);
    if (! inputContext)
    {
        return std::nullopt;
    }
    const auto releaseContext = wil::scope_exit([&] { ImmReleaseContext(bridgeEdit, inputContext); });

    CANDIDATEFORM candidateForm{};
    if (ImmGetCandidateWindow(inputContext, index, &candidateForm) == FALSE)
    {
        return std::nullopt;
    }

    return candidateForm;
}

class SingleCellGridModel final : public RedSalamander::DxUi::IDxGridModel
{
public:
    explicit SingleCellGridModel(RedSalamander::DxUi::GridCellData cellData) : _cellData(std::move(cellData))
    {
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return 1u;
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return 1u;
    }

    [[nodiscard]] RedSalamander::DxUi::GridColumnDesc GetColumn(size_t /*columnIndex*/) const override
    {
        RedSalamander::DxUi::GridColumnDesc column;
        column.id       = L"status";
        column.title    = L"Status";
        column.widthDip = 160.0f;
        return column;
    }

    void GetCellData(size_t /*rowIndex*/, size_t /*columnIndex*/, RedSalamander::DxUi::GridCellData& outCell) const override
    {
        outCell = _cellData;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        return rowId == 0u ? std::optional<size_t>(0u) : std::nullopt;
    }

private:
    RedSalamander::DxUi::GridCellData _cellData;
};

class MultiRowGridModel final : public RedSalamander::DxUi::IDxGridModel
{
public:
    explicit MultiRowGridModel(size_t rowCount) : _rowCount(rowCount)
    {
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rowCount;
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return 1u;
    }

    [[nodiscard]] RedSalamander::DxUi::GridColumnDesc GetColumn(size_t /*columnIndex*/) const override
    {
        RedSalamander::DxUi::GridColumnDesc column;
        column.id       = L"name";
        column.title    = L"Name";
        column.widthDip = 180.0f;
        return column;
    }

    void GetCellData(size_t rowIndex, size_t /*columnIndex*/, RedSalamander::DxUi::GridCellData& outCell) const override
    {
        outCell.kind = RedSalamander::DxUi::GridCellKind::Text;
        outCell.text = std::format(L"Row {:02}", rowIndex);
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (rowId >= _rowCount)
        {
            return std::nullopt;
        }

        return static_cast<size_t>(rowId);
    }

private:
    size_t _rowCount = 0u;
};

class GroupedGridModel final : public RedSalamander::DxUi::IDxGridModel
{
public:
    struct Group
    {
        uint64_t stableId = 0u;
        std::wstring title;
        size_t startRowIndex = 0u;
        size_t rowCount      = 0u;
        bool collapsed       = false;
    };

    explicit GroupedGridModel(size_t rowCount) : _rowCount(rowCount)
    {
    }

    void SetGroups(std::vector<Group> groups)
    {
        _groups = std::move(groups);
    }

    bool SetGroupCollapsed(uint64_t stableId, bool collapsed)
    {
        for (Group& group : _groups)
        {
            if (group.stableId == stableId)
            {
                group.collapsed = collapsed;
                return true;
            }
        }

        return false;
    }

    [[nodiscard]] bool IsGroupCollapsed(uint64_t stableId) const
    {
        for (const Group& group : _groups)
        {
            if (group.stableId == stableId)
            {
                return group.collapsed;
            }
        }

        return false;
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rowCount;
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return 1u;
    }

    [[nodiscard]] RedSalamander::DxUi::GridColumnDesc GetColumn(size_t /*columnIndex*/) const override
    {
        RedSalamander::DxUi::GridColumnDesc column;
        column.id       = L"name";
        column.title    = L"Name";
        column.widthDip = 180.0f;
        return column;
    }

    void GetCellData(size_t rowIndex, size_t /*columnIndex*/, RedSalamander::DxUi::GridCellData& outCell) const override
    {
        outCell.kind = RedSalamander::DxUi::GridCellKind::Text;
        outCell.text = std::format(L"Row {:02}", rowIndex);
    }

    [[nodiscard]] size_t GetGroupCount() const noexcept override
    {
        return _groups.size();
    }

    [[nodiscard]] RedSalamander::DxUi::GridGroupDesc GetGroup(size_t groupIndex) const override
    {
        const Group& group = _groups.at(groupIndex);
        return RedSalamander::DxUi::GridGroupDesc{
            .stableId      = group.stableId,
            .title         = group.title,
            .startRowIndex = group.startRowIndex,
            .rowCount      = group.rowCount,
            .collapsed     = group.collapsed,
        };
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (rowId >= _rowCount)
        {
            return std::nullopt;
        }

        return static_cast<size_t>(rowId);
    }

private:
    size_t _rowCount = 0u;
    std::vector<Group> _groups;
};

class LargeGridModel final : public RedSalamander::DxUi::IDxGridModel
{
public:
    LargeGridModel(size_t rowCount, size_t columnCount, float columnWidthDip) : _rowCount(rowCount), _columnCount(columnCount), _columnWidthDip(columnWidthDip)
    {
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rowCount;
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columnCount;
    }

    [[nodiscard]] RedSalamander::DxUi::GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        RedSalamander::DxUi::GridColumnDesc column;
        column.id       = std::format(L"col-{}", columnIndex);
        column.title    = std::format(L"Column {}", columnIndex);
        column.widthDip = _columnWidthDip;
        return column;
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, RedSalamander::DxUi::GridCellData& outCell) const override
    {
        outCell.kind = RedSalamander::DxUi::GridCellKind::Text;
        outCell.text = std::format(L"R{}C{}", rowIndex, columnIndex);
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (rowId >= _rowCount)
        {
            return std::nullopt;
        }

        return static_cast<size_t>(rowId);
    }

private:
    size_t _rowCount      = 0u;
    size_t _columnCount   = 0u;
    float _columnWidthDip = 120.0f;
};

class MutableRowGridModel final : public RedSalamander::DxUi::IDxGridModel
{
public:
    void SetRowIds(std::vector<uint64_t> rowIds)
    {
        _rowIds = std::move(rowIds);
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rowIds.size();
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return 1u;
    }

    [[nodiscard]] RedSalamander::DxUi::GridColumnDesc GetColumn(size_t /*columnIndex*/) const override
    {
        RedSalamander::DxUi::GridColumnDesc column;
        column.id       = L"name";
        column.title    = L"Name";
        column.widthDip = 180.0f;
        return column;
    }

    void GetCellData(size_t rowIndex, size_t /*columnIndex*/, RedSalamander::DxUi::GridCellData& outCell) const override
    {
        outCell.kind = RedSalamander::DxUi::GridCellKind::Text;
        outCell.text = std::format(L"Row {}", _rowIds.at(rowIndex));
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return _rowIds.at(rowIndex);
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        for (size_t rowIndex = 0; rowIndex < _rowIds.size(); ++rowIndex)
        {
            if (_rowIds[rowIndex] == rowId)
            {
                return rowIndex;
            }
        }

        return std::nullopt;
    }

private:
    std::vector<uint64_t> _rowIds;
};

class ColumnLayoutGridModel final : public RedSalamander::DxUi::IDxGridModel
{
public:
    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return 1u;
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return 3u;
    }

    [[nodiscard]] RedSalamander::DxUi::GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        using namespace RedSalamander::DxUi;

        switch (columnIndex)
        {
            case 0u:
                return GridColumnDesc{
                    .id            = L"name",
                    .title         = L"Name",
                    .widthDip      = 160.0f,
                    .minWidthDip   = 80.0f,
                    .kind          = GridColumnKind::Text,
                    .sortable      = true,
                    .multiline     = false,
                    .textAlignment = DWRITE_TEXT_ALIGNMENT_LEADING,
                };
            case 1u:
                return GridColumnDesc{
                    .id            = L"path",
                    .title         = L"Path",
                    .widthDip      = 220.0f,
                    .minWidthDip   = 100.0f,
                    .kind          = GridColumnKind::Text,
                    .sortable      = true,
                    .multiline     = false,
                    .textAlignment = DWRITE_TEXT_ALIGNMENT_LEADING,
                };
            default:
                return GridColumnDesc{
                    .id            = L"modified",
                    .title         = L"Modified",
                    .widthDip      = 180.0f,
                    .minWidthDip   = 96.0f,
                    .kind          = GridColumnKind::Text,
                    .sortable      = true,
                    .multiline     = false,
                    .textAlignment = DWRITE_TEXT_ALIGNMENT_LEADING,
                };
        }
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, RedSalamander::DxUi::GridCellData& outCell) const override
    {
        Require(rowIndex == 0u, "column-layout model uses one row");
        outCell.kind      = RedSalamander::DxUi::GridCellKind::Text;
        outCell.multiline = false;

        switch (columnIndex)
        {
            case 0u: outCell.text = L"alpha.txt"; return;
            case 1u: outCell.text = L"C:\\Data"; return;
            default: outCell.text = L"2026-03-15"; return;
        }
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return static_cast<uint64_t>(rowIndex + 1u);
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (rowId == 0u || rowId > GetRowCount())
        {
            return std::nullopt;
        }

        return static_cast<size_t>(rowId - 1u);
    }
};

class RecordingGridDelegate : public RedSalamander::DxUi::IDxGridDelegate
{
public:
    using RedSalamander::DxUi::IDxGridDelegate::OnGridCheckboxToggled;
    using RedSalamander::DxUi::IDxGridDelegate::OnGridContextMenu;
    using RedSalamander::DxUi::IDxGridDelegate::OnGridGroupToggled;
    using RedSalamander::DxUi::IDxGridDelegate::OnGridRowActivated;
    using RedSalamander::DxUi::IDxGridDelegate::OnGridSelectionChanged;

    void OnGridSortRequested(const RedSalamander::DxUi::GridSortSpec& sortSpec) override
    {
        ++sortRequestedCount;
        lastSortSpec = sortSpec;
    }

    void OnGridSelectionChanged(RedSalamander::DxUi::Grid& sender) override
    {
        lastSelectionSender = &sender;
        ++selectionChangedCount;
    }

    void OnGridRowActivated(RedSalamander::DxUi::Grid& sender, size_t rowIndex) override
    {
        lastActivatedSender = &sender;
        ++rowActivatedCount;
        lastActivatedRow = rowIndex;
    }

    void OnGridContextMenu(size_t rowIndex, POINT screenPoint) override
    {
        ++contextMenuCount;
        lastContextMenuRow   = rowIndex;
        lastContextMenuPoint = screenPoint;
    }

    void OnGridGroupToggled(uint64_t groupStableId, bool collapsed) override
    {
        ++groupToggleCount;
        lastGroupStableId  = groupStableId;
        lastGroupCollapsed = collapsed;
    }

    size_t sortRequestedCount = 0u;
    RedSalamander::DxUi::GridSortSpec lastSortSpec{};
    RedSalamander::DxUi::Grid* lastSelectionSender = nullptr;
    size_t selectionChangedCount                   = 0u;
    RedSalamander::DxUi::Grid* lastActivatedSender = nullptr;
    size_t rowActivatedCount                       = 0u;
    size_t lastActivatedRow                        = 0u;
    size_t contextMenuCount                        = 0u;
    size_t lastContextMenuRow                      = 0u;
    POINT lastContextMenuPoint{};
    size_t groupToggleCount    = 0u;
    uint64_t lastGroupStableId = 0u;
    bool lastGroupCollapsed    = false;
};

class ModelSwappingGridDelegate final : public RecordingGridDelegate
{
public:
    using RecordingGridDelegate::OnGridSelectionChanged;

    void OnGridSelectionChanged(RedSalamander::DxUi::Grid& sender) override
    {
        RecordingGridDelegate::OnGridSelectionChanged(sender);
        sender.SetModel(nullptr);
    }
};

class CollapsibleGroupedGridDelegate final : public RecordingGridDelegate
{
public:
    using RecordingGridDelegate::OnGridGroupToggled;

    explicit CollapsibleGroupedGridDelegate(GroupedGridModel& model) : _model(&model)
    {
    }

    void OnGridGroupToggled(uint64_t groupStableId, bool collapsed) override
    {
        RecordingGridDelegate::OnGridGroupToggled(groupStableId, collapsed);
        Require(_model->SetGroupCollapsed(groupStableId, collapsed), "grouped grid delegate updates the requested group collapse state");
    }

private:
    GroupedGridModel* _model = nullptr;
};

class CheckboxGridModel final : public RedSalamander::DxUi::IDxGridModel
{
public:
    struct Row
    {
        std::wstring label;
        bool checked = false;
        bool enabled = true;
    };

    explicit CheckboxGridModel(size_t checkboxColumnIndex) : _checkboxColumnIndex(checkboxColumnIndex)
    {
    }

    void SetRows(std::vector<Row> rows)
    {
        _rows = std::move(rows);
    }

    [[nodiscard]] bool SetChecked(size_t rowIndex, size_t columnIndex, bool checked)
    {
        if (rowIndex >= _rows.size() || columnIndex != _checkboxColumnIndex)
        {
            return false;
        }
        _rows[rowIndex].checked = checked;
        return true;
    }

    [[nodiscard]] bool IsChecked(size_t rowIndex) const
    {
        return _rows.at(rowIndex).checked;
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rows.size();
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return 2u;
    }

    [[nodiscard]] RedSalamander::DxUi::GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        RedSalamander::DxUi::GridColumnDesc column;
        if (columnIndex == _checkboxColumnIndex)
        {
            column.id        = L"enabled";
            column.title     = L"Enabled";
            column.widthDip  = 120.0f;
            column.multiline = false;
            return column;
        }

        column.id        = L"name";
        column.title     = L"Name";
        column.widthDip  = 180.0f;
        column.multiline = false;
        return column;
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, RedSalamander::DxUi::GridCellData& outCell) const override
    {
        const Row& row = _rows.at(rowIndex);
        if (columnIndex == _checkboxColumnIndex)
        {
            outCell.kind      = RedSalamander::DxUi::GridCellKind::Checkbox;
            outCell.text      = L"Enabled";
            outCell.checked   = row.checked;
            outCell.enabled   = row.enabled;
            outCell.multiline = false;
            return;
        }

        outCell.kind      = RedSalamander::DxUi::GridCellKind::Text;
        outCell.text      = row.label;
        outCell.multiline = false;
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return static_cast<uint64_t>(rowIndex + 1u);
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (rowId == 0u || rowId > _rows.size())
        {
            return std::nullopt;
        }

        return static_cast<size_t>(rowId - 1u);
    }

private:
    std::vector<Row> _rows;
    size_t _checkboxColumnIndex = 0u;
};

class LargeCheckboxGridModel final : public RedSalamander::DxUi::IDxGridModel
{
public:
    LargeCheckboxGridModel(size_t rowCount, size_t checkboxColumnIndex) : _rowCount(rowCount), _checkboxColumnIndex(checkboxColumnIndex)
    {
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rowCount;
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return 2u;
    }

    [[nodiscard]] RedSalamander::DxUi::GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        RedSalamander::DxUi::GridColumnDesc column;
        if (columnIndex == _checkboxColumnIndex)
        {
            column.id        = L"enabled";
            column.title     = L"Enabled";
            column.widthDip  = 120.0f;
            column.multiline = false;
            return column;
        }

        column.id        = L"name";
        column.title     = L"Name";
        column.widthDip  = 180.0f;
        column.multiline = false;
        return column;
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, RedSalamander::DxUi::GridCellData& outCell) const override
    {
        if (columnIndex == _checkboxColumnIndex)
        {
            outCell.kind      = RedSalamander::DxUi::GridCellKind::Checkbox;
            outCell.text      = L"Enabled";
            outCell.checked   = (rowIndex % 2u) == 0u;
            outCell.enabled   = true;
            outCell.multiline = false;
            return;
        }

        outCell.kind      = RedSalamander::DxUi::GridCellKind::Text;
        outCell.text      = std::format(L"Rule {:05}", rowIndex);
        outCell.multiline = false;
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return static_cast<uint64_t>(rowIndex + 1u);
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (rowId == 0u || rowId > _rowCount)
        {
            return std::nullopt;
        }

        return static_cast<size_t>(rowId - 1u);
    }

private:
    size_t _rowCount            = 0u;
    size_t _checkboxColumnIndex = 0u;
};

class RecordingCheckboxGridDelegate final : public RecordingGridDelegate
{
public:
    using RecordingGridDelegate::OnGridCheckboxToggled;

    explicit RecordingCheckboxGridDelegate(CheckboxGridModel& model) : _model(&model)
    {
    }

    void OnGridCheckboxToggled(size_t rowIndex, size_t columnIndex, bool checked) override
    {
        ++toggleCount;
        lastToggleRow     = rowIndex;
        lastToggleColumn  = columnIndex;
        lastToggleChecked = checked;
        Require(_model->SetChecked(rowIndex, columnIndex, checked), "checkbox grid delegate updates the requested checkbox cell");
    }

    size_t toggleCount      = 0u;
    size_t lastToggleRow    = 0u;
    size_t lastToggleColumn = 0u;
    bool lastToggleChecked  = false;

private:
    CheckboxGridModel* _model = nullptr;
};

class DedicatedCheckboxColumnGridModel final : public RedSalamander::DxUi::IDxGridModel
{
public:
    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return 1u;
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return 2u;
    }

    [[nodiscard]] RedSalamander::DxUi::GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        RedSalamander::DxUi::GridColumnDesc column;
        if (columnIndex == 0u)
        {
            column.id            = L"enabled";
            column.title         = L"";
            column.kind          = RedSalamander::DxUi::GridColumnKind::Checkbox;
            column.widthDip      = 36.0f;
            column.minWidthDip   = 36.0f;
            column.sortable      = false;
            column.multiline     = false;
            column.textAlignment = DWRITE_TEXT_ALIGNMENT_CENTER;
            return column;
        }

        column.id        = L"name";
        column.title     = L"Name";
        column.widthDip  = 180.0f;
        column.multiline = false;
        return column;
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, RedSalamander::DxUi::GridCellData& outCell) const override
    {
        Require(rowIndex == 0u, "dedicated checkbox model uses one row");
        if (columnIndex == 0u)
        {
            outCell.kind = RedSalamander::DxUi::GridCellKind::Checkbox;
            outCell.text.clear();
            outCell.checked   = _checked;
            outCell.enabled   = true;
            outCell.multiline = false;
            return;
        }

        outCell.kind      = RedSalamander::DxUi::GridCellKind::Text;
        outCell.text      = L"Alpha";
        outCell.multiline = false;
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return static_cast<uint64_t>(rowIndex + 1u);
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        return rowId == 1u ? std::optional<size_t>(0u) : std::nullopt;
    }

    void SetChecked(bool checked) noexcept
    {
        _checked = checked;
    }

    [[nodiscard]] bool IsChecked() const noexcept
    {
        return _checked;
    }

private:
    bool _checked = false;
};

class DedicatedCheckboxGridDelegate final : public RecordingGridDelegate
{
public:
    using RecordingGridDelegate::OnGridCheckboxToggled;

    explicit DedicatedCheckboxGridDelegate(DedicatedCheckboxColumnGridModel& model) : _model(&model)
    {
    }

    void OnGridCheckboxToggled(size_t rowIndex, size_t columnIndex, bool checked) override
    {
        ++toggleCount;
        lastToggleRow     = rowIndex;
        lastToggleColumn  = columnIndex;
        lastToggleChecked = checked;
        Require(rowIndex == 0u && columnIndex == 0u, "dedicated checkbox toggle targets the dedicated column");
        _model->SetChecked(checked);
    }

    size_t toggleCount      = 0u;
    size_t lastToggleRow    = 0u;
    size_t lastToggleColumn = 0u;
    bool lastToggleChecked  = false;

private:
    DedicatedCheckboxColumnGridModel* _model = nullptr;
};

class StateImageColumnGridModel final : public RedSalamander::DxUi::IDxGridModel
{
public:
    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return 1u;
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return 2u;
    }

    [[nodiscard]] RedSalamander::DxUi::GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        RedSalamander::DxUi::GridColumnDesc column;
        if (columnIndex == 0u)
        {
            column.id            = L"state";
            column.title         = L"";
            column.kind          = RedSalamander::DxUi::GridColumnKind::StateImage;
            column.widthDip      = 40.0f;
            column.minWidthDip   = 40.0f;
            column.sortable      = false;
            column.multiline     = false;
            column.textAlignment = DWRITE_TEXT_ALIGNMENT_CENTER;
            return column;
        }

        column.id        = L"name";
        column.title     = L"Name";
        column.widthDip  = 180.0f;
        column.multiline = false;
        return column;
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, RedSalamander::DxUi::GridCellData& outCell) const override
    {
        Require(rowIndex == 0u, "state-image model uses one row");
        if (columnIndex == 0u)
        {
            outCell.kind = RedSalamander::DxUi::GridCellKind::IconText;
            outCell.text.clear();
            outCell.iconText  = L"!";
            outCell.multiline = false;
            return;
        }

        outCell.kind      = RedSalamander::DxUi::GridCellKind::Text;
        outCell.text      = L"Warning";
        outCell.multiline = false;
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return static_cast<uint64_t>(rowIndex + 1u);
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        return rowId == 1u ? std::optional<size_t>(0u) : std::nullopt;
    }
};

class MutableTreeModel final : public RedSalamander::DxUi::IDxTreeModel
{
public:
    void SetVisibleItems(std::vector<RedSalamander::DxUi::TreeItemData> items)
    {
        _items = std::move(items);
    }

    [[nodiscard]] size_t GetVisibleItemCount() const noexcept override
    {
        return _items.size();
    }

    void GetVisibleItem(size_t visibleIndex, RedSalamander::DxUi::TreeItemData& outItem) const override
    {
        outItem = _items.at(visibleIndex);
    }

private:
    std::vector<RedSalamander::DxUi::TreeItemData> _items;
};

class RecordingTreeDelegate final : public RedSalamander::DxUi::IDxTreeDelegate
{
public:
    void OnTreeSelectionChanged(uint64_t itemId) override
    {
        ++selectionChangedCount;
        lastSelectedItemId = itemId;
    }

    void OnTreeItemInvoked(uint64_t itemId) override
    {
        ++invokedCount;
        lastInvokedItemId = itemId;
    }

    void OnTreeToggleExpanded(uint64_t itemId, bool expanded) override
    {
        ++toggleCount;
        lastToggledItemId = itemId;
        lastExpandedState = expanded;
    }

    void OnTreeContextMenu(uint64_t itemId, POINT screenPoint) override
    {
        ++contextMenuCount;
        lastContextMenuItemId = itemId;
        lastContextMenuPoint  = screenPoint;
    }

    size_t selectionChangedCount   = 0u;
    size_t invokedCount            = 0u;
    size_t toggleCount             = 0u;
    size_t contextMenuCount        = 0u;
    uint64_t lastSelectedItemId    = 0u;
    uint64_t lastInvokedItemId     = 0u;
    uint64_t lastToggledItemId     = 0u;
    uint64_t lastContextMenuItemId = 0u;
    bool lastExpandedState         = false;
    POINT lastContextMenuPoint{};
};

class PaintTraceControl final : public RedSalamander::DxUi::Control
{
public:
    PaintTraceControl(std::vector<std::string>& events, std::string paintTag, std::string overlayTag = {})
        : _events(&events),
          _paintTag(std::move(paintTag)),
          _overlayTag(std::move(overlayTag))
    {
    }

    void Paint(RedSalamander::DxUi::WindowHost& /*host*/) const override
    {
        _events->push_back(_paintTag);
    }

    void PaintOverlay(RedSalamander::DxUi::WindowHost& /*host*/) const override
    {
        if (! _overlayTag.empty())
        {
            _events->push_back(_overlayTag);
        }
    }

private:
    std::vector<std::string>* _events = nullptr;
    std::string _paintTag;
    std::string _overlayTag;
};

class PaintTraceComboBox final : public RedSalamander::DxUi::ComboBox
{
public:
    explicit PaintTraceComboBox(std::vector<std::string>& events) : _events(&events)
    {
    }

    void Paint(RedSalamander::DxUi::WindowHost& /*host*/) const override
    {
        _events->push_back("combo-base");
    }

    void PaintOverlay(RedSalamander::DxUi::WindowHost& /*host*/) const override
    {
        if (GetHitBounds().bottom > GetBounds().bottom)
        {
            _events->push_back("combo-popup");
        }
    }

private:
    std::vector<std::string>* _events = nullptr;
};

class EmptyGridModel final : public RedSalamander::DxUi::IDxGridModel
{
public:
    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return 0u;
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return 1u;
    }

    [[nodiscard]] RedSalamander::DxUi::GridColumnDesc GetColumn(size_t /*columnIndex*/) const override
    {
        RedSalamander::DxUi::GridColumnDesc column;
        column.id       = L"empty";
        column.title    = L"Empty";
        column.widthDip = 160.0f;
        return column;
    }

    void GetCellData(size_t /*rowIndex*/, size_t /*columnIndex*/, RedSalamander::DxUi::GridCellData& /*outCell*/) const override
    {
        ++cellAccessCount;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t /*rowId*/) const noexcept override
    {
        return std::nullopt;
    }

    mutable size_t cellAccessCount = 0u;
};

struct TrackingControlState
{
    size_t focusGainCount  = 0u;
    size_t focusLossCount  = 0u;
    size_t hoverEnterCount = 0u;
    size_t hoverLeaveCount = 0u;
    size_t mouseDownCount  = 0u;
    size_t mouseMoveCount  = 0u;
    size_t mouseLeaveCount = 0u;
    size_t mouseUpCount    = 0u;
    size_t mouseWheelCount = 0u;
};

class TrackingControl final : public RedSalamander::DxUi::Control
{
public:
    explicit TrackingControl(TrackingControlState& state) : _state(&state)
    {
        SetFocusable(true);
    }

    void Paint(RedSalamander::DxUi::WindowHost& /*host*/) const override
    {
    }

    bool OnMouseMove(RedSalamander::DxUi::WindowHost& /*host*/, D2D1_POINT_2F /*point*/, UINT /*modifiers*/) override
    {
        ++_state->mouseMoveCount;
        return true;
    }

    bool OnMouseLeave(RedSalamander::DxUi::WindowHost& host) override
    {
        ++_state->mouseLeaveCount;
        return Control::OnMouseLeave(host);
    }

    bool OnMouseDown(RedSalamander::DxUi::WindowHost& /*host*/, D2D1_POINT_2F /*point*/, bool rightButton, UINT /*modifiers*/) override
    {
        ++_state->mouseDownCount;
        return ! rightButton;
    }

    bool OnMouseUp(RedSalamander::DxUi::WindowHost& /*host*/, D2D1_POINT_2F /*point*/, bool /*rightButton*/, UINT /*modifiers*/) override
    {
        ++_state->mouseUpCount;
        return true;
    }

    bool OnMouseWheel(RedSalamander::DxUi::WindowHost& /*host*/, D2D1_POINT_2F /*point*/, float /*wheelDelta*/, UINT /*modifiers*/) override
    {
        ++_state->mouseWheelCount;
        return true;
    }

protected:
    void OnFocusChanged(RedSalamander::DxUi::WindowHost& host, bool focused) override
    {
        if (focused)
        {
            ++_state->focusGainCount;
        }
        else
        {
            ++_state->focusLossCount;
        }
        Control::OnFocusChanged(host, focused);
    }

    void OnHoverChanged(RedSalamander::DxUi::WindowHost& host, bool hovered) override
    {
        if (hovered)
        {
            ++_state->hoverEnterCount;
        }
        else
        {
            ++_state->hoverLeaveCount;
        }
        Control::OnHoverChanged(host, hovered);
    }

private:
    TrackingControlState* _state = nullptr;
};

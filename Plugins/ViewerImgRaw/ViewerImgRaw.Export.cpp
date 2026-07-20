#include "ViewerImgRaw.h"

#include "ViewerImgRaw.Internal.h"
#include "ViewerImgRaw.ResourcePolicy.h"

#include <algorithm>
#include <array>
#include <format>
#include <functional>
#include <limits>
#include <new>
#include <optional>
#include <string>
#include <string_view>
#include <thread>
#include <vector>

#include <objbase.h>
#include <oleauto.h>
#include <propidl.h>
#include <shobjidl_core.h>
#include <wincodec.h>

#include "Helpers.h"
#include "PathUtils.h"
#include "resource.h"

namespace
{
static const int kViewerImgRawModuleAnchor = 0;

[[nodiscard]] std::wstring FormatExportError(const UINT resourceId, const HRESULT hr)
{
    return FormatStringResource(g_hInstance, resourceId, static_cast<uint32_t>(hr));
}

std::wstring SanitizeFileName(std::wstring_view name) noexcept
{
    std::wstring out;
    out.reserve(name.size());

    for (const wchar_t ch : name)
    {
        if (ch < 0x20)
        {
            out.push_back(L'_');
            continue;
        }

        switch (ch)
        {
            case L'<':
            case L'>':
            case L':':
            case L'"':
            case L'/':
            case L'\\':
            case L'|':
            case L'?':
            case L'*': out.push_back(L'_'); break;
            default: out.push_back(ch); break;
        }
    }

    while (! out.empty() && (out.back() == L' ' || out.back() == L'.'))
    {
        out.pop_back();
    }

    if (out.empty())
    {
        out = L"image";
    }

    return out;
}

enum class ExportFormat : uint8_t
{
    Png  = 0,
    Jpeg = 1,
    Tiff = 2,
    Bmp  = 3,
    Gif  = 4,
    Wmp  = 5,
};

struct ExportSaveDialogResult final
{
    std::wstring path;
    ExportFormat formatFromFilter = ExportFormat::Png;
};

std::optional<ExportFormat> ExportFormatFromExtension(std::wstring_view ext) noexcept
{
    if (EqualsIgnoreCase(ext, L".jpg") || EqualsIgnoreCase(ext, L".jpeg"))
    {
        return ExportFormat::Jpeg;
    }
    if (EqualsIgnoreCase(ext, L".png"))
    {
        return ExportFormat::Png;
    }
    if (EqualsIgnoreCase(ext, L".tif") || EqualsIgnoreCase(ext, L".tiff"))
    {
        return ExportFormat::Tiff;
    }
    if (EqualsIgnoreCase(ext, L".bmp") || EqualsIgnoreCase(ext, L".dib"))
    {
        return ExportFormat::Bmp;
    }
    if (EqualsIgnoreCase(ext, L".gif"))
    {
        return ExportFormat::Gif;
    }
    if (EqualsIgnoreCase(ext, L".jxr") || EqualsIgnoreCase(ext, L".wdp") || EqualsIgnoreCase(ext, L".hdp"))
    {
        return ExportFormat::Wmp;
    }

    return std::nullopt;
}

std::wstring_view DefaultExtensionForExportFormat(ExportFormat format) noexcept
{
    switch (format)
    {
        case ExportFormat::Png: return L"png";
        case ExportFormat::Jpeg: return L"jpg";
        case ExportFormat::Tiff: return L"tif";
        case ExportFormat::Bmp: return L"bmp";
        case ExportFormat::Gif: return L"gif";
        case ExportFormat::Wmp: return L"jxr";
        default: return L"png";
    }
}

UINT FilterIndexForExportFormat(ExportFormat format) noexcept
{
    // 1-based indices
    switch (format)
    {
        case ExportFormat::Png: return 1;
        case ExportFormat::Jpeg: return 2;
        case ExportFormat::Tiff: return 3;
        case ExportFormat::Bmp: return 4;
        case ExportFormat::Gif: return 5;
        case ExportFormat::Wmp: return 6;
        default: return 1;
    }
}

ExportFormat ExportFormatFromFilterIndex(UINT index) noexcept
{
    switch (index)
    {
        case 2: return ExportFormat::Jpeg;
        case 3: return ExportFormat::Tiff;
        case 4: return ExportFormat::Bmp;
        case 5: return ExportFormat::Gif;
        case 6: return ExportFormat::Wmp;
        case 1:
        default: return ExportFormat::Png;
    }
}

std::optional<ExportSaveDialogResult> ShowExportSaveDialog(HWND hwnd, ExportFormat defaultFormat, std::wstring_view suggestedFileName) noexcept
{
    const HRESULT coinitHr  = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    const bool shouldUninit = SUCCEEDED(coinitHr);
    auto coUninit           = wil::scope_exit([&]
    {
        if (shouldUninit)
        {
            CoUninitialize();
        }
    });

    // If the host initialized COM with a different apartment model, don't crash; proceed best-effort.
    if (FAILED(coinitHr) && coinitHr != RPC_E_CHANGED_MODE)
    {
        return std::nullopt;
    }

    wil::com_ptr<IFileSaveDialog> dialog;
    const HRESULT hr = CoCreateInstance(CLSID_FileSaveDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(dialog.put()));
    if (FAILED(hr) || ! dialog)
    {
        return std::nullopt;
    }

    const std::wstring title = LoadStringResource(g_hInstance, IDS_VIEWERRAW_DIALOG_EXPORT_TITLE);
    if (! title.empty())
    {
        static_cast<void>(dialog->SetTitle(title.c_str()));
    }

    DWORD options = 0;
    static_cast<void>(dialog->GetOptions(&options));
    options |= FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST | FOS_OVERWRITEPROMPT;
    static_cast<void>(dialog->SetOptions(options));

    if (! suggestedFileName.empty())
    {
        const std::wstring fileName = SanitizeFileName(suggestedFileName);
        static_cast<void>(dialog->SetFileName(fileName.c_str()));
    }

    const std::wstring filterPng  = LoadStringResource(g_hInstance, IDS_VIEWERRAW_DIALOG_FILTER_PNG);
    const std::wstring filterJpeg = LoadStringResource(g_hInstance, IDS_VIEWERRAW_DIALOG_FILTER_JPEG);
    const std::wstring filterTiff = LoadStringResource(g_hInstance, IDS_VIEWERRAW_DIALOG_FILTER_TIFF);
    const std::wstring filterBmp  = LoadStringResource(g_hInstance, IDS_VIEWERRAW_DIALOG_FILTER_BMP);
    const std::wstring filterGif  = LoadStringResource(g_hInstance, IDS_VIEWERRAW_DIALOG_FILTER_GIF);
    const std::wstring filterJxr  = LoadStringResource(g_hInstance, IDS_VIEWERRAW_DIALOG_FILTER_JXR);

    const std::array<COMDLG_FILTERSPEC, 6> specs{
        COMDLG_FILTERSPEC{filterPng.c_str(), L"*.png"},
        COMDLG_FILTERSPEC{filterJpeg.c_str(), L"*.jpg;*.jpeg"},
        COMDLG_FILTERSPEC{filterTiff.c_str(), L"*.tif;*.tiff"},
        COMDLG_FILTERSPEC{filterBmp.c_str(), L"*.bmp;*.dib"},
        COMDLG_FILTERSPEC{filterGif.c_str(), L"*.gif"},
        COMDLG_FILTERSPEC{filterJxr.c_str(), L"*.jxr;*.wdp;*.hdp"},
    };

    static_cast<void>(dialog->SetFileTypes(static_cast<UINT>(specs.size()), specs.data()));
    static_cast<void>(dialog->SetFileTypeIndex(FilterIndexForExportFormat(defaultFormat)));
    static_cast<void>(dialog->SetDefaultExtension(DefaultExtensionForExportFormat(defaultFormat).data()));

    const HRESULT showHr = dialog->Show(hwnd);
    if (FAILED(showHr))
    {
        return std::nullopt;
    }

    UINT typeIndex = 1;
    static_cast<void>(dialog->GetFileTypeIndex(&typeIndex));

    wil::com_ptr<IShellItem> item;
    const HRESULT itemHr = dialog->GetResult(item.put());
    if (FAILED(itemHr) || ! item)
    {
        return std::nullopt;
    }

    wil::unique_cotaskmem_string path;
    const HRESULT nameHr = item->GetDisplayName(SIGDN_FILESYSPATH, path.put());
    if (FAILED(nameHr) || ! path)
    {
        return std::nullopt;
    }

    ExportSaveDialogResult result{};
    result.path             = path.get();
    result.formatFromFilter = ExportFormatFromFilterIndex(typeIndex);
    return result;
}

struct ExportEncoderOptions final
{
    uint32_t jpegQualityPercent  = 90;
    uint32_t jpegSubsampling     = 0; // WICJpegYCrCbSubsamplingOption
    uint32_t pngFilter           = 0; // WICPngFilterOption
    bool pngInterlace            = false;
    uint32_t tiffCompression     = 0; // WICTiffCompressionOption
    bool bmpUseV5Header32bppBGRA = true;
    bool gifInterlace            = false;
    uint32_t wmpQualityPercent   = 90;
    bool wmpLossless             = false;
};

HRESULT TryWriteEncoderOptionBool(IPropertyBag2* options, const wchar_t* name, bool value) noexcept
{
    if (! options || ! name || name[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    PROPBAG2 prop{};
    prop.pstrName = const_cast<LPOLESTR>(name);

    VARIANT v{};
    VariantInit(&v);
    v.vt      = VT_BOOL;
    v.boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;

    return options->Write(1, &prop, &v);
}

HRESULT TryWriteEncoderOptionUInt(IPropertyBag2* options, const wchar_t* name, uint32_t value) noexcept
{
    if (! options || ! name || name[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    PROPBAG2 prop{};
    prop.pstrName = const_cast<LPOLESTR>(name);

    VARIANT v{};
    VariantInit(&v);
    v.vt    = VT_UI4;
    v.ulVal = value;

    return options->Write(1, &prop, &v);
}

HRESULT TryWriteEncoderOptionFloat(IPropertyBag2* options, const wchar_t* name, float value) noexcept
{
    if (! options || ! name || name[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    PROPBAG2 prop{};
    prop.pstrName = const_cast<LPOLESTR>(name);

    VARIANT v{};
    VariantInit(&v);
    v.vt     = VT_R4;
    v.fltVal = value;

    return options->Write(1, &prop, &v);
}

void ApplyExportEncoderOptions(IPropertyBag2* options, ExportFormat exportFormat, const ExportEncoderOptions& cfg) noexcept
{
    if (! options)
    {
        return;
    }

    switch (exportFormat)
    {
        case ExportFormat::Jpeg:
        {
            const float q = std::clamp(static_cast<float>(cfg.jpegQualityPercent) / 100.0f, 0.0f, 1.0f);
            static_cast<void>(TryWriteEncoderOptionFloat(options, L"ImageQuality", q));
            static_cast<void>(TryWriteEncoderOptionUInt(options, L"JpegYCrCbSubsampling", cfg.jpegSubsampling));
            return;
        }
        case ExportFormat::Png:
            static_cast<void>(TryWriteEncoderOptionUInt(options, L"FilterOption", cfg.pngFilter));
            static_cast<void>(TryWriteEncoderOptionBool(options, L"InterlaceOption", cfg.pngInterlace));
            return;
        case ExportFormat::Tiff: static_cast<void>(TryWriteEncoderOptionUInt(options, L"TiffCompressionMethod", cfg.tiffCompression)); return;
        case ExportFormat::Bmp: static_cast<void>(TryWriteEncoderOptionBool(options, L"EnableV5Header32bppBGRA", cfg.bmpUseV5Header32bppBGRA)); return;
        case ExportFormat::Wmp:
        {
            const float q = std::clamp(static_cast<float>(cfg.wmpQualityPercent) / 100.0f, 0.0f, 1.0f);
            static_cast<void>(TryWriteEncoderOptionFloat(options, L"ImageQuality", q));
            static_cast<void>(TryWriteEncoderOptionBool(options, L"Lossless", cfg.wmpLossless));
            return;
        }
        case ExportFormat::Gif: static_cast<void>(TryWriteEncoderOptionBool(options, L"InterlaceOption", cfg.gifInterlace)); return;
        default: return;
    }
}

HRESULT EncodeBgraToImageFileWic(const std::wstring& outputPath,
                                 ExportFormat exportFormat,
                                 uint32_t width,
                                 uint32_t height,
                                 const std::vector<uint8_t>& bgra,
                                 const ExportEncoderOptions& exportOptions,
                                 std::wstring& outStatusMessage) noexcept
{
    outStatusMessage.clear();

    if (outputPath.empty() || width == 0 || height == 0 || bgra.empty())
    {
        return E_INVALIDARG;
    }

    ViewerImgRawResource::DecodedImageLayout layout{};
    const ViewerImgRawResource::ValidationError validation =
        ViewerImgRawResource::ValidateDecodedImage(width, height, ViewerImgRawResource::kProductionDecodedImagePolicy, layout);
    if (validation != ViewerImgRawResource::ValidationError::None || layout.bgraBytes > std::numeric_limits<UINT>::max() ||
        layout.rowBytes > std::numeric_limits<UINT>::max())
    {
        outStatusMessage = LoadStringResource(g_hInstance, IDS_VIEWERRAW_EXPORT_IMAGE_TOO_LARGE);
        return validation == ViewerImgRawResource::ValidationError::InvalidDimensions ? E_INVALIDARG : HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }
    if (layout.bgraBytes > bgra.size())
    {
        outStatusMessage = LoadStringResource(g_hInstance, IDS_VIEWERRAW_EXPORT_BUFFER_TRUNCATED);
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    // Atomic export: encode into a sibling temp file in the destination's own directory, then
    // rename over the destination only after a fully successful encode. Any failure (disk full,
    // device removed, encoder error, GIF palette quantization failure, oversized dimensions, ...)
    // must never destroy the original the user confirmed overwriting. Mirrors ViewerWeb's
    // temp-file + MoveFileExW export pattern. The directory is derived with std::wstring (not
    // std::filesystem) to keep this noexcept function's existing exception profile — only
    // std::bad_alloc, as with the std::format calls below. Extended Win32 paths preserve the
    // sibling-temp guarantee for destinations beyond MAX_PATH.
    const std::wstring outputWin32Path = Common::Paths::ToExtendedWin32Path(outputPath);
    std::wstring tempPath;
    wil::unique_hfile reservation;
    const Common::Paths::UniqueSiblingFileOptions options{.prefix = L".rsi-", .suffix = L".tmp", .maximumAttempts = 16u};
    const HRESULT tempHr = Common::Paths::CreateUniqueSiblingFile(outputWin32Path, options, tempPath, reservation);
    if (FAILED(tempHr))
    {
        outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_TEMP_CREATE_FAILED_FMT, tempHr);
        return tempHr;
    }
    reservation.reset();

    // The CREATE_NEW reservation already created a zero-byte temp file. Arm cleanup BEFORE the
    // WIC objects are created so that on every early-return failure path the WIC stream (declared
    // below) is destroyed first — closing its file handle — and this scope_exit fires last,
    // deleting the temp after the handle is gone. On success, `committed` is set only after the
    // rename, so the (already renamed-away) temp is never deleted.
    bool committed   = false;
    auto cleanupTemp = wil::scope_exit([&tempPath, &committed]() noexcept
    {
        if (! committed)
        {
            static_cast<void>(DeleteFileW(tempPath.c_str())); // best-effort; the original destination is untouched
        }
    });

    wil::com_ptr<IWICImagingFactory> factory;
    HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(factory.put()));
    if (FAILED(hr) || ! factory)
    {
        outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_WIC_FACTORY_FAILED_FMT, hr);
        return FAILED(hr) ? hr : E_FAIL;
    }

    wil::com_ptr<IWICStream> stream;
    hr = factory->CreateStream(stream.put());
    if (FAILED(hr) || ! stream)
    {
        outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_WIC_STREAM_FAILED_FMT, hr);
        return FAILED(hr) ? hr : E_FAIL;
    }

    hr = stream->InitializeFromFilename(tempPath.c_str(), GENERIC_WRITE);
    if (FAILED(hr))
    {
        outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_FILE_OPEN_FAILED_FMT, hr);
        return hr;
    }

    GUID container                 = GUID_ContainerFormatPng;
    WICPixelFormatGUID pixelFormat = GUID_WICPixelFormat32bppBGRA;
    switch (exportFormat)
    {
        case ExportFormat::Png: container = GUID_ContainerFormatPng; break;
        case ExportFormat::Jpeg:
            container   = GUID_ContainerFormatJpeg;
            pixelFormat = GUID_WICPixelFormat24bppBGR;
            break;
        case ExportFormat::Tiff: container = GUID_ContainerFormatTiff; break;
        case ExportFormat::Bmp: container = GUID_ContainerFormatBmp; break;
        case ExportFormat::Gif:
            container   = GUID_ContainerFormatGif;
            pixelFormat = GUID_WICPixelFormat8bppIndexed;
            break;
        case ExportFormat::Wmp:
            container   = GUID_ContainerFormatWmp;
            pixelFormat = GUID_WICPixelFormat24bppBGR;
            break;
        default: outStatusMessage = LoadStringResource(g_hInstance, IDS_VIEWERRAW_EXPORT_FORMAT_UNSUPPORTED); return E_INVALIDARG;
    }

    wil::com_ptr<IWICBitmapEncoder> encoder;
    hr = factory->CreateEncoder(container, nullptr, encoder.put());
    if (FAILED(hr) || ! encoder)
    {
        outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_ENCODER_CREATE_FAILED_FMT, hr);
        return FAILED(hr) ? hr : E_FAIL;
    }

    hr = encoder->Initialize(stream.get(), WICBitmapEncoderNoCache);
    if (FAILED(hr))
    {
        outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_ENCODER_INIT_FAILED_FMT, hr);
        return hr;
    }

    wil::com_ptr<IWICBitmapFrameEncode> frame;
    wil::com_ptr<IPropertyBag2> frameOptions;
    hr = encoder->CreateNewFrame(frame.put(), frameOptions.put());
    if (FAILED(hr) || ! frame)
    {
        outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_FRAME_CREATE_FAILED_FMT, hr);
        return FAILED(hr) ? hr : E_FAIL;
    }

    ApplyExportEncoderOptions(frameOptions.get(), exportFormat, exportOptions);
    hr = frame->Initialize(frameOptions.get());
    if (FAILED(hr))
    {
        outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_FRAME_INIT_FAILED_FMT, hr);
        return hr;
    }

    hr = frame->SetSize(width, height);
    if (FAILED(hr))
    {
        outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_SIZE_FAILED_FMT, hr);
        return hr;
    }

    WICPixelFormatGUID format = pixelFormat;
    hr                        = frame->SetPixelFormat(&format);
    if (FAILED(hr))
    {
        outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_PIXEL_FORMAT_FAILED_FMT, hr);
        return hr;
    }

    const UINT stride     = static_cast<UINT>(layout.rowBytes);
    const UINT bufferSize = static_cast<UINT>(layout.bgraBytes);

    wil::com_ptr<IWICBitmap> bitmap;
    hr = factory->CreateBitmapFromMemory(width, height, GUID_WICPixelFormat32bppBGRA, stride, bufferSize, const_cast<BYTE*>(bgra.data()), bitmap.put());
    if (FAILED(hr) || ! bitmap)
    {
        outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_BITMAP_CREATE_FAILED_FMT, hr);
        return FAILED(hr) ? hr : E_FAIL;
    }

    wil::com_ptr<IWICBitmapSource> source = bitmap;
    wil::com_ptr<IWICPalette> paletteToSet;
    if (format == GUID_WICPixelFormat8bppIndexed)
    {
        wil::com_ptr<IWICPalette> palette;
        hr = factory->CreatePalette(palette.put());
        if (FAILED(hr) || ! palette)
        {
            outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_PALETTE_CREATE_FAILED_FMT, hr);
            return FAILED(hr) ? hr : E_FAIL;
        }

        hr = palette->InitializeFromBitmap(bitmap.get(), 256, FALSE);
        if (FAILED(hr))
        {
            outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_PALETTE_INIT_FAILED_FMT, hr);
            return hr;
        }

        wil::com_ptr<IWICFormatConverter> converter;
        hr = factory->CreateFormatConverter(converter.put());
        if (FAILED(hr) || ! converter)
        {
            outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_CONVERTER_CREATE_FAILED_FMT, hr);
            return FAILED(hr) ? hr : E_FAIL;
        }

        hr = converter->Initialize(bitmap.get(), format, WICBitmapDitherTypeErrorDiffusion, palette.get(), 0.0, WICBitmapPaletteTypeCustom);
        if (FAILED(hr))
        {
            outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_INDEXED_CONVERT_FAILED_FMT, hr);
            return hr;
        }

        source       = converter;
        paletteToSet = palette;
    }
    else if (format != GUID_WICPixelFormat32bppBGRA)
    {
        wil::com_ptr<IWICFormatConverter> converter;
        hr = factory->CreateFormatConverter(converter.put());
        if (FAILED(hr) || ! converter)
        {
            outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_CONVERTER_CREATE_FAILED_FMT, hr);
            return FAILED(hr) ? hr : E_FAIL;
        }

        hr = converter->Initialize(bitmap.get(), format, WICBitmapDitherTypeNone, nullptr, 0.0, WICBitmapPaletteTypeCustom);
        if (FAILED(hr))
        {
            outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_TARGET_CONVERT_FAILED_FMT, hr);
            return hr;
        }

        source = converter;
    }

    if (paletteToSet)
    {
        static_cast<void>(frame->SetPalette(paletteToSet.get()));
    }

    hr = frame->WriteSource(source.get(), nullptr);
    if (FAILED(hr))
    {
        outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_FRAME_WRITE_FAILED_FMT, hr);
        return hr;
    }

    hr = frame->Commit();
    if (FAILED(hr))
    {
        outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_FRAME_COMMIT_FAILED_FMT, hr);
        return hr;
    }

    hr = encoder->Commit();
    if (FAILED(hr))
    {
        outStatusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_ENCODER_COMMIT_FAILED_FMT, hr);
        return hr;
    }

    // encoder->Commit() flushed all encoded bytes into the stream, but the WIC encoder still holds
    // an AddRef on the IWICStream (taken in encoder->Initialize(stream.get(), ...)), and the frame
    // holds the encoder, so the temp file's OS handle stays open. Release the entire WIC object
    // graph that pins the stream BEFORE the rename, or MoveFileExW fails with a sharing violation.
    frame.reset();
    encoder.reset();
    source.reset();
    bitmap.reset();
    paletteToSet.reset();
    stream.reset();

    if (MoveFileExW(tempPath.c_str(), outputWin32Path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) == 0)
    {
        const DWORD lastError = GetLastError();
        outStatusMessage      = FormatExportError(IDS_VIEWERRAW_EXPORT_FINALIZE_FAILED_FMT, HRESULT_FROM_WIN32(lastError));
        return HRESULT_FROM_WIN32(lastError); // cleanupTemp removes the temp; the original is untouched
    }

    committed = true; // the temp has been renamed onto the destination — do not delete it
    return S_OK;
}
} // namespace

void ViewerImgRaw::BeginExport(HWND hwnd)
{
    BeginExportImpl(hwnd, nullptr);
}

void ViewerImgRaw::BeginExportImpl(HWND hwnd, const wchar_t* outputPathOverride)
{
    if (! hwnd)
    {
        return;
    }

    if (! HasDisplayImage() || ! _currentImage)
    {
        if (_hostAlerts)
        {
            const std::wstring message = LoadStringResource(g_hInstance, IDS_VIEWERRAW_EXPORT_NO_IMAGE);
            HostAlertRequest req{};
            req.version      = 1;
            req.sizeBytes    = sizeof(req);
            req.scope        = HOST_ALERT_SCOPE_WINDOW;
            req.modality     = HOST_ALERT_MODELESS;
            req.severity     = HOST_ALERT_INFO;
            req.targetWindow = _hWnd.get();
            req.title        = _metaName.empty() ? nullptr : _metaName.c_str();
            req.message      = message.c_str();
            req.closable     = TRUE;
            static_cast<void>(_hostAlerts->ShowAlert(&req, nullptr));
            _alertVisible = true;
        }
        return;
    }

    uint32_t w = 0;
    uint32_t h = 0;
    std::vector<uint8_t> pixels; // owned snapshot — survives the modal save dialog's nested message pump
    const bool exportingThumb = IsDisplayingThumbnail();
    const CachedImage* image  = _currentImage; // UI-thread-affine; non-null here per the guard at the top of BeginExport
    if (image != nullptr)                      // defensive re-check; no message pump runs between that guard and here
    {
        w      = exportingThumb ? image->thumbWidth : image->rawWidth;
        h      = exportingThumb ? image->thumbHeight : image->rawHeight;
        pixels = exportingThumb ? image->thumbBgra : image->rawBgra; // deep copy BEFORE the modal dialog pumps
    }

    if (w == 0 || h == 0 || pixels.empty())
    {
        if (_hostAlerts)
        {
            const std::wstring message = LoadStringResource(g_hInstance, IDS_VIEWERRAW_EXPORT_NO_IMAGE);
            HostAlertRequest req{};
            req.version      = 1;
            req.sizeBytes    = sizeof(req);
            req.scope        = HOST_ALERT_SCOPE_WINDOW;
            req.modality     = HOST_ALERT_MODELESS;
            req.severity     = HOST_ALERT_INFO;
            req.targetWindow = _hWnd.get();
            req.title        = _metaName.empty() ? nullptr : _metaName.c_str();
            req.message      = message.c_str();
            req.closable     = TRUE;
            static_cast<void>(_hostAlerts->ShowAlert(&req, nullptr));
            _alertVisible = true;
        }
        return;
    }

    std::wstring baseName = LeafNameFromPath(_currentPath);
    if (baseName.empty())
    {
        baseName = L"image";
    }

    if (const size_t dot = baseName.find_last_of(L'.'); dot != std::wstring::npos && dot != 0)
    {
        baseName = baseName.substr(0, dot);
    }

    ExportFormat defaultFormat = ExportFormat::Png;
    if (const std::optional<ExportFormat> fromExt = ExportFormatFromExtension(PathExtensionView(_currentPath)))
    {
        defaultFormat = *fromExt;
    }

    std::wstring suggested;
    suggested = baseName;
    suggested.push_back(L'.');
    suggested.append(DefaultExtensionForExportFormat(defaultFormat).data());

    std::optional<ExportSaveDialogResult> save;
    if (outputPathOverride && outputPathOverride[0] != L'\0')
    {
        ExportSaveDialogResult overridden{};
        overridden.path             = outputPathOverride;
        overridden.formatFromFilter = ExportFormatFromExtension(PathExtensionView(overridden.path)).value_or(defaultFormat);
        save                        = std::move(overridden);
    }
    else
    {
        save = ShowExportSaveDialog(hwnd, defaultFormat, suggested);
    }
    if (! save || save->path.empty())
    {
        return;
    }

    std::wstring output       = save->path;
    ExportFormat exportFormat = save->formatFromFilter;

    const std::wstring_view ext = PathExtensionView(output);
    if (! ext.empty())
    {
        if (const std::optional<ExportFormat> fromExt = ExportFormatFromExtension(ext))
        {
            exportFormat = *fromExt;
        }
        else
        {
            if (_hostAlerts)
            {
                const std::wstring message = LoadStringResource(g_hInstance, IDS_VIEWERRAW_EXPORT_UNSUPPORTED_EXTENSION);
                HostAlertRequest req{};
                req.version      = 1;
                req.sizeBytes    = sizeof(req);
                req.scope        = HOST_ALERT_SCOPE_WINDOW;
                req.modality     = HOST_ALERT_MODELESS;
                req.severity     = HOST_ALERT_WARNING;
                req.targetWindow = _hWnd.get();
                req.title        = _metaName.empty() ? nullptr : _metaName.c_str();
                req.message      = message.c_str();
                req.closable     = TRUE;
                static_cast<void>(_hostAlerts->ShowAlert(&req, nullptr));
                _alertVisible = true;
            }
            return;
        }
    }
    else
    {
        output.push_back(L'.');
        output.append(DefaultExtensionForExportFormat(exportFormat).data());
    }

    ExportEncoderOptions encoderOptions{};
    encoderOptions.jpegQualityPercent      = std::clamp(_config.exportJpegQualityPercent, 1u, 100u);
    encoderOptions.jpegSubsampling         = std::clamp(_config.exportJpegSubsampling, 0u, 4u);
    encoderOptions.pngFilter               = std::clamp(_config.exportPngFilter, 0u, 6u);
    encoderOptions.pngInterlace            = _config.exportPngInterlace;
    encoderOptions.tiffCompression         = std::clamp(_config.exportTiffCompression, 0u, 7u);
    encoderOptions.bmpUseV5Header32bppBGRA = _config.exportBmpUseV5Header32bppBGRA;
    encoderOptions.gifInterlace            = _config.exportGifInterlace;
    encoderOptions.wmpQualityPercent       = std::clamp(_config.exportWmpQualityPercent, 1u, 100u);
    encoderOptions.wmpLossless             = _config.exportWmpLossless;

    AddRef();
    struct AsyncExportWorkItem final
    {
        AsyncExportWorkItem()                                      = default;
        AsyncExportWorkItem(const AsyncExportWorkItem&)            = delete;
        AsyncExportWorkItem& operator=(const AsyncExportWorkItem&) = delete;

        wil::unique_hmodule moduleKeepAlive;
        std::function<void()> work;
    };

    auto ctx = std::unique_ptr<AsyncExportWorkItem>(new (std::nothrow) AsyncExportWorkItem{});
    if (! ctx)
    {
        Release(); // balance the AddRef() above
        return;
    }

    ctx->moduleKeepAlive = AcquireModuleReferenceFromAddress(&kViewerImgRawModuleAnchor);
    if (! ctx->moduleKeepAlive)
    {
        Debug::Error(L"ViewerImgRaw: Failed to pin the plugin module for export work");
        Release(); // balance the AddRef() above
        return;
    }
    ctx->work = [this, hwnd, exportFormat, width = w, height = h, encoderOptions, output = std::move(output), pixels = std::move(pixels)]() mutable
    {
        auto releaseSelf = wil::scope_exit([&] { Release(); });

        const HRESULT coinitHr  = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        const bool shouldUninit = SUCCEEDED(coinitHr);
        auto coUninit           = wil::scope_exit([&]
        {
            if (shouldUninit)
            {
                CoUninitialize();
            }
        });

        std::unique_ptr<AsyncExportResult> result(new (std::nothrow) AsyncExportResult{});
        if (! result)
        {
            return;
        }

        result->viewer     = this;
        result->outputPath = output;

        if (FAILED(coinitHr) && coinitHr != RPC_E_CHANGED_MODE)
        {
            result->hr            = coinitHr;
            result->statusMessage = FormatExportError(IDS_VIEWERRAW_EXPORT_COM_INIT_FAILED_FMT, coinitHr);
        }
        else
        {
            result->hr = EncodeBgraToImageFileWic(output, exportFormat, width, height, pixels, encoderOptions, result->statusMessage);
        }

        if (! hwnd || GetWindowLongPtrW(hwnd, GWLP_USERDATA) != reinterpret_cast<LONG_PTR>(this))
        {
            return;
        }

        static_cast<void>(PostMessagePayload(hwnd, kAsyncExportCompleteMessage, 0, std::move(result)));
    };

    NotifyViewerImgRawAsyncWorkQueued();
    const BOOL queued = TrySubmitThreadpoolCallback(
        [](PTP_CALLBACK_INSTANCE instance, void* context) noexcept
    {
        std::unique_ptr<AsyncExportWorkItem> ctx(static_cast<AsyncExportWorkItem*>(context));
        if (! ctx)
        {
            NotifyViewerImgRawAsyncWorkCompleted();
            return;
        }

        if (ctx->moduleKeepAlive)
        {
            TransferModulePinToCallbackReturn(instance, ctx->moduleKeepAlive);
        }
        auto completeAsyncWork = wil::scope_exit([]() noexcept { NotifyViewerImgRawAsyncWorkCompleted(); });
        if (ctx->work)
        {
            ctx->work();
        }
    },
        ctx.get(),
        nullptr);

    if (queued == 0)
    {
        NotifyViewerImgRawAsyncWorkCompleted();
        Debug::Error(L"ViewerImgRaw: Failed to queue export work item");
        Release(); // balance the AddRef(); ctx's destructor frees the work item
        return;
    }

    ctx.release(); // success only — the threadpool now owns the work item
}

void ViewerImgRaw::OnAsyncExportComplete(std::unique_ptr<AsyncExportResult> result) noexcept
{
    if (! result)
    {
        return;
    }

    if (_hostAlerts)
    {
        std::wstring message;
        HostAlertSeverity severity = HOST_ALERT_INFO;
        if (SUCCEEDED(result->hr))
        {
            message  = FormatStringResource(g_hInstance, IDS_VIEWERRAW_EXPORT_SUCCESS_FMT, result->outputPath);
            severity = HOST_ALERT_INFO;
        }
        else
        {
            message  = result->statusMessage.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERRAW_EXPORT_FAILED) : result->statusMessage;
            severity = HOST_ALERT_WARNING;
        }

        HostAlertRequest req{};
        req.version      = 1;
        req.sizeBytes    = sizeof(req);
        req.scope        = HOST_ALERT_SCOPE_WINDOW;
        req.modality     = HOST_ALERT_MODELESS;
        req.severity     = severity;
        req.targetWindow = _hWnd.get();
        req.title        = _metaName.empty() ? nullptr : _metaName.c_str();
        req.message      = message.c_str();
        req.closable     = TRUE;
        static_cast<void>(_hostAlerts->ShowAlert(&req, nullptr));
        _alertVisible = true;
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
}

#include "ViewerImgRaw.h"

#include "ViewerImgRaw.Internal.h"

#include <algorithm>
#include <cmath>
#include <cstring>
#include <ctime>
#include <cwctype>
#include <format>
#include <functional>
#include <iterator>
#include <limits>
#include <new>
#include <thread>

#include <d2d1.h>
#include <libraw/libraw.h>
#include <turbojpeg.h>
#include <wincodec.h>

#include "Helpers.h"
#include "resource.h"

namespace
{
static const int kViewerImgRawModuleAnchor = 0;

std::wstring WideFromUtf8(std::string_view text)
{
    if (text.empty())
    {
        return {};
    }

    const int srcLen = static_cast<int>(std::min<size_t>(text.size(), static_cast<size_t>(std::numeric_limits<int>::max())));

    int len       = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), srcLen, nullptr, 0);
    UINT codePage = CP_UTF8;
    DWORD flags   = MB_ERR_INVALID_CHARS;
    if (len <= 0)
    {
        codePage = CP_ACP;
        flags    = 0;
        len      = MultiByteToWideChar(codePage, flags, text.data(), srcLen, nullptr, 0);
    }

    if (len <= 0)
    {
        return {};
    }

    std::wstring out(static_cast<size_t>(len), L'\0');
    const int converted = MultiByteToWideChar(codePage, flags, text.data(), srcLen, out.data(), len);
    if (converted <= 0)
    {
        return {};
    }

    return out;
}

struct ExifData final
{
    std::wstring camera;
    std::wstring lens;
    std::wstring dateTime;
    float iso            = 0.0f;
    float shutterSeconds = 0.0f;
    float aperture       = 0.0f;
    float focalLengthMm  = 0.0f;
    uint16_t orientation = 1; // EXIF orientation (1..8)
    bool valid           = false;
};

std::wstring TrimSpaces(std::wstring_view text);

[[nodiscard]] bool InRange(size_t offset, size_t length, size_t size) noexcept
{
    return offset <= size && length <= (size - offset);
}

[[nodiscard]] uint16_t ReadU16(const uint8_t* p, bool littleEndian) noexcept
{
    if (littleEndian)
    {
        return static_cast<uint16_t>(static_cast<uint16_t>(p[0]) | (static_cast<uint16_t>(p[1]) << 8));
    }
    return static_cast<uint16_t>((static_cast<uint16_t>(p[0]) << 8) | static_cast<uint16_t>(p[1]));
}

[[nodiscard]] uint32_t ReadU32(const uint8_t* p, bool littleEndian) noexcept
{
    if (littleEndian)
    {
        return static_cast<uint32_t>(static_cast<uint32_t>(p[0]) | (static_cast<uint32_t>(p[1]) << 8) | (static_cast<uint32_t>(p[2]) << 16) |
                                     (static_cast<uint32_t>(p[3]) << 24));
    }
    return static_cast<uint32_t>((static_cast<uint32_t>(p[0]) << 24) | (static_cast<uint32_t>(p[1]) << 16) | (static_cast<uint32_t>(p[2]) << 8) |
                                 static_cast<uint32_t>(p[3]));
}

[[nodiscard]] uint16_t ClampExifOrientation(uint16_t orientation) noexcept
{
    return (orientation >= 1 && orientation <= 8) ? orientation : static_cast<uint16_t>(1);
}

[[nodiscard]] std::wstring ReadTiffAscii(const uint8_t* tiffBase, size_t tiffSize, const uint8_t* valueBytes, uint32_t valueOrOffset, uint32_t count) noexcept
{
    if (! tiffBase || tiffSize == 0 || count == 0)
    {
        return {};
    }

    std::string bytes;
    if (count <= 4)
    {
        bytes.assign(reinterpret_cast<const char*>(valueBytes), reinterpret_cast<const char*>(valueBytes) + count);
    }
    else
    {
        const uint32_t offset = valueOrOffset;
        if (! InRange(offset, count, tiffSize))
        {
            return {};
        }

        bytes.assign(reinterpret_cast<const char*>(tiffBase + offset), reinterpret_cast<const char*>(tiffBase + offset + count));
    }

    const size_t nul = bytes.find('\0');
    if (nul != std::string::npos)
    {
        bytes.resize(nul);
    }

    return TrimSpaces(WideFromUtf8(bytes));
}

[[nodiscard]] bool ReadTiffShort(
    const uint8_t* tiffBase, size_t tiffSize, bool littleEndian, const uint8_t* valueBytes, uint32_t valueOrOffset, uint32_t count, uint16_t& outValue) noexcept
{
    outValue = 0;
    if (! tiffBase || tiffSize == 0 || count == 0)
    {
        return false;
    }

    if (count * 2u <= 4u)
    {
        if (! valueBytes)
        {
            return false;
        }
        outValue = ReadU16(valueBytes, littleEndian);
        return true;
    }

    const uint32_t offset = valueOrOffset;
    if (! InRange(offset, 2, tiffSize))
    {
        return false;
    }

    outValue = ReadU16(tiffBase + offset, littleEndian);
    return true;
}

[[nodiscard]] bool
ReadTiffLong(const uint8_t* tiffBase, size_t tiffSize, bool littleEndian, uint32_t valueOrOffset, uint32_t count, uint32_t& outValue) noexcept
{
    outValue = 0;
    if (! tiffBase || tiffSize == 0 || count == 0)
    {
        return false;
    }

    if (count * 4u <= 4u)
    {
        outValue = valueOrOffset;
        return true;
    }

    const uint32_t offset = valueOrOffset;
    if (! InRange(offset, 4, tiffSize))
    {
        return false;
    }

    outValue = ReadU32(tiffBase + offset, littleEndian);
    return true;
}

[[nodiscard]] bool
ReadTiffRational(const uint8_t* tiffBase, size_t tiffSize, bool littleEndian, uint32_t valueOrOffset, uint32_t count, float& outValue) noexcept
{
    outValue = 0.0f;
    if (! tiffBase || tiffSize == 0 || count == 0)
    {
        return false;
    }

    const uint32_t offset = valueOrOffset;
    if (! InRange(offset, 8, tiffSize))
    {
        return false;
    }

    const uint32_t numer = ReadU32(tiffBase + offset, littleEndian);
    const uint32_t denom = ReadU32(tiffBase + offset + 4, littleEndian);
    if (denom == 0)
    {
        return false;
    }

    outValue = static_cast<float>(static_cast<double>(numer) / static_cast<double>(denom));
    return true;
}

ExifData ExtractExifFromJpeg(const uint8_t* data, size_t sizeBytes) noexcept
{
    ExifData out{};

    if (! data || sizeBytes < 4)
    {
        return out;
    }

    if (data[0] != 0xFF || data[1] != 0xD8)
    {
        return out;
    }

    size_t pos = 2;
    while (pos + 4 <= sizeBytes)
    {
        if (data[pos] != 0xFF)
        {
            ++pos;
            continue;
        }

        while (pos < sizeBytes && data[pos] == 0xFF)
        {
            ++pos;
        }
        if (pos >= sizeBytes)
        {
            break;
        }

        const uint8_t marker = data[pos++];
        if (marker == 0xDA || marker == 0xD9) // SOS or EOI
        {
            break;
        }

        if (pos + 2 > sizeBytes)
        {
            break;
        }

        const uint16_t segLen = static_cast<uint16_t>((static_cast<uint16_t>(data[pos]) << 8) | static_cast<uint16_t>(data[pos + 1]));
        pos += 2;
        if (segLen < 2)
        {
            break;
        }

        const size_t segDataLen = static_cast<size_t>(segLen - 2);
        if (! InRange(pos, segDataLen, sizeBytes))
        {
            break;
        }

        if (marker == 0xE1 && segDataLen >= 6)
        {
            const uint8_t* seg = data + pos;
            if (segDataLen >= 14 && std::memcmp(seg, "Exif\0\0", 6) == 0)
            {
                const uint8_t* tiff   = seg + 6;
                const size_t tiffSize = segDataLen - 6;

                if (tiffSize >= 8)
                {
                    const bool little = (tiff[0] == 'I' && tiff[1] == 'I') ? true : ((tiff[0] == 'M' && tiff[1] == 'M') ? false : true);

                    const uint16_t magic = ReadU16(tiff + 2, little);
                    if (magic == 42)
                    {
                        const uint32_t ifd0Off = ReadU32(tiff + 4, little);
                        if (ifd0Off != 0 && InRange(ifd0Off, 2, tiffSize))
                        {
                            std::wstring make;
                            std::wstring model;
                            std::wstring dateTime;
                            std::wstring dateTimeOriginal;
                            std::wstring lensModel;
                            uint16_t orientation = 1;
                            float iso            = 0.0f;
                            float shutter        = 0.0f;
                            float aperture       = 0.0f;
                            float focal          = 0.0f;

                            uint32_t exifIfdOff = 0;

                            const uint16_t count      = ReadU16(tiff + ifd0Off, little);
                            const size_t entriesStart = static_cast<size_t>(ifd0Off) + 2;
                            for (uint16_t i = 0; i < count; ++i)
                            {
                                const size_t entryOff = entriesStart + static_cast<size_t>(i) * 12u;
                                if (! InRange(entryOff, 12, tiffSize))
                                {
                                    break;
                                }

                                const uint16_t tag                   = ReadU16(tiff + entryOff, little);
                                [[maybe_unused]] const uint16_t type = ReadU16(tiff + entryOff + 2, little);
                                const uint32_t cnt                   = ReadU32(tiff + entryOff + 4, little);
                                const uint8_t* valueBytes            = tiff + entryOff + 8;
                                const uint32_t valueOrOffset         = ReadU32(valueBytes, little);

                                switch (tag)
                                {
                                    case 0x0112:
                                    {
                                        uint16_t v = 0;
                                        if (ReadTiffShort(tiff, tiffSize, little, valueBytes, valueOrOffset, cnt, v))
                                        {
                                            orientation = ClampExifOrientation(v);
                                        }
                                        break;
                                    }
                                    case 0x010F: make = ReadTiffAscii(tiff, tiffSize, valueBytes, valueOrOffset, cnt); break;
                                    case 0x0110: model = ReadTiffAscii(tiff, tiffSize, valueBytes, valueOrOffset, cnt); break;
                                    case 0x0132: dateTime = ReadTiffAscii(tiff, tiffSize, valueBytes, valueOrOffset, cnt); break;
                                    case 0x8769: static_cast<void>(ReadTiffLong(tiff, tiffSize, little, valueOrOffset, cnt, exifIfdOff)); break;
                                    default: break;
                                }
                            }

                            if (exifIfdOff != 0 && InRange(exifIfdOff, 2, tiffSize))
                            {
                                const uint16_t exifCount      = ReadU16(tiff + exifIfdOff, little);
                                const size_t exifEntriesStart = static_cast<size_t>(exifIfdOff) + 2;
                                for (uint16_t i = 0; i < exifCount; ++i)
                                {
                                    const size_t entryOff = exifEntriesStart + static_cast<size_t>(i) * 12u;
                                    if (! InRange(entryOff, 12, tiffSize))
                                    {
                                        break;
                                    }

                                    const uint16_t tag           = ReadU16(tiff + entryOff, little);
                                    const uint16_t type          = ReadU16(tiff + entryOff + 2, little);
                                    const uint32_t cnt           = ReadU32(tiff + entryOff + 4, little);
                                    const uint8_t* valueBytes    = tiff + entryOff + 8;
                                    const uint32_t valueOrOffset = ReadU32(valueBytes, little);

                                    switch (tag)
                                    {
                                        case 0x9003: dateTimeOriginal = ReadTiffAscii(tiff, tiffSize, valueBytes, valueOrOffset, cnt); break;
                                        case 0xA434: lensModel = ReadTiffAscii(tiff, tiffSize, valueBytes, valueOrOffset, cnt); break;
                                        case 0x8827:
                                        {
                                            uint16_t isoShort = 0;
                                            if (type == 3 && ReadTiffShort(tiff, tiffSize, little, valueBytes, valueOrOffset, cnt, isoShort))
                                            {
                                                iso = static_cast<float>(isoShort);
                                            }
                                            else
                                            {
                                                uint32_t isoLong = 0;
                                                if (type == 4 && ReadTiffLong(tiff, tiffSize, little, valueOrOffset, cnt, isoLong))
                                                {
                                                    iso = static_cast<float>(isoLong);
                                                }
                                            }
                                            break;
                                        }
                                        case 0x8833:
                                        {
                                            uint16_t isoShort = 0;
                                            if (type == 3 && ReadTiffShort(tiff, tiffSize, little, valueBytes, valueOrOffset, cnt, isoShort))
                                            {
                                                iso = static_cast<float>(isoShort);
                                                break;
                                            }

                                            uint32_t isoLong = 0;
                                            if (type == 4 && ReadTiffLong(tiff, tiffSize, little, valueOrOffset, cnt, isoLong))
                                            {
                                                iso = static_cast<float>(isoLong);
                                            }
                                            break;
                                        }
                                        case 0x829A:
                                            if (type == 5)
                                            {
                                                static_cast<void>(ReadTiffRational(tiff, tiffSize, little, valueOrOffset, cnt, shutter));
                                            }
                                            break;
                                        case 0x829D:
                                            if (type == 5)
                                            {
                                                static_cast<void>(ReadTiffRational(tiff, tiffSize, little, valueOrOffset, cnt, aperture));
                                            }
                                            break;
                                        case 0x920A:
                                            if (type == 5)
                                            {
                                                static_cast<void>(ReadTiffRational(tiff, tiffSize, little, valueOrOffset, cnt, focal));
                                            }
                                            break;
                                        default: break;
                                    }
                                }
                            }

                            make  = TrimSpaces(make);
                            model = TrimSpaces(model);
                            if (! make.empty() && ! model.empty())
                            {
                                out.camera = std::format(L"{} {}", make, model);
                            }
                            else
                            {
                                out.camera = ! model.empty() ? model : make;
                            }
                            out.lens           = TrimSpaces(lensModel);
                            out.dateTime       = TrimSpaces(! dateTimeOriginal.empty() ? dateTimeOriginal : dateTime);
                            out.iso            = iso;
                            out.shutterSeconds = shutter;
                            out.aperture       = aperture;
                            out.focalLengthMm  = focal;
                            out.orientation    = orientation;
                            out.valid = ! out.camera.empty() || ! out.lens.empty() || ! out.dateTime.empty() || out.iso > 0.0f || out.shutterSeconds > 0.0f ||
                                        out.aperture > 0.0f || out.focalLengthMm > 0.0f || out.orientation != 1;
                        }
                    }
                }

                break;
            }
        }

        pos += segDataLen;
    }

    out.orientation = ClampExifOrientation(out.orientation);
    return out;
}

std::wstring TrimSpaces(std::wstring_view text)
{
    size_t start = 0;
    while (start < text.size() && std::iswspace(text[start]) != 0)
    {
        ++start;
    }

    size_t end = text.size();
    while (end > start && std::iswspace(text[end - 1]) != 0)
    {
        --end;
    }

    if (start == 0 && end == text.size())
    {
        return std::wstring(text);
    }

    return std::wstring(text.substr(start, end - start));
}

ExifData ExtractExifData(const LibRaw& raw) noexcept
{
    ExifData exif{};

    const std::wstring make  = TrimSpaces(WideFromUtf8(raw.imgdata.idata.make));
    const std::wstring model = TrimSpaces(WideFromUtf8(raw.imgdata.idata.model));
    if (! make.empty() && ! model.empty())
    {
        exif.camera = std::format(L"{} {}", make, model);
    }
    else
    {
        exif.camera = ! model.empty() ? model : make;
    }

    exif.lens = TrimSpaces(WideFromUtf8(raw.imgdata.lens.Lens));

    exif.iso            = raw.imgdata.other.iso_speed;
    exif.shutterSeconds = raw.imgdata.other.shutter;
    exif.aperture       = raw.imgdata.other.aperture;
    exif.focalLengthMm  = raw.imgdata.other.focal_len;

    const time_t ts = raw.imgdata.other.timestamp;
    if (ts > 0)
    {
        std::tm tm{};
        if (localtime_s(&tm, &ts) == 0)
        {
            wchar_t buffer[64]{};
            if (std::wcsftime(buffer, std::size(buffer), L"%Y-%m-%d %H:%M:%S", &tm) > 0)
            {
                exif.dateTime = buffer;
            }
        }
    }

    exif.valid = ! exif.camera.empty() || ! exif.lens.empty() || ! exif.dateTime.empty() || exif.iso > 0.0f || exif.shutterSeconds > 0.0f ||
                 exif.aperture > 0.0f || exif.focalLengthMm > 0.0f;
    return exif;
}

struct RawDecodeSettings final
{
    bool halfSize    = false;
    bool useCameraWb = true;
    bool autoWb      = false;
};

struct LibRawProgressHost final
{
    const std::atomic_uint64_t* requestIdCounter = nullptr;
    uint64_t requestId                           = 0;
    HWND hwnd                                    = nullptr;
};

struct LibRawProgressContext final
{
    LibRawProgressHost host{};
    int lastPercent = -1;
    int lastStage   = -1;
};

int LibRawProgressCallback(void* data, enum LibRaw_progress stage, int iteration, int expected) noexcept
{
    auto* ctx = static_cast<LibRawProgressContext*>(data);
    if (! ctx || ! ctx->host.requestIdCounter || ! ctx->host.hwnd)
    {
        return 0;
    }

    if (ctx->host.requestIdCounter->load(std::memory_order_acquire) != ctx->host.requestId)
    {
        return 1;
    }

    int percent = -1;
    if (expected > 0 && iteration >= 0)
    {
        const long long numer = static_cast<long long>(iteration) * 100ll;
        const long long denom = static_cast<long long>(expected);
        if (denom > 0)
        {
            percent = static_cast<int>(std::clamp<long long>(numer / denom, 0ll, 100ll));
        }
    }

    const int stageInt        = static_cast<int>(stage);
    const bool stageChanged   = stageInt != ctx->lastStage;
    const bool percentChanged = percent != ctx->lastPercent;
    if (stageChanged || percentChanged)
    {
        ctx->lastStage   = stageInt;
        ctx->lastPercent = percent;
        static_cast<void>(PostMessageW(ctx->host.hwnd, kAsyncProgressMessage, static_cast<WPARAM>(stageInt), static_cast<LPARAM>(percent)));
    }

    return 0;
}

HRESULT ReadFileAllBytes(IFileSystem* fileSystem, const std::wstring& path, std::vector<uint8_t>& outBytes, std::wstring& outStatusMessage) noexcept
{
    outBytes.clear();
    outStatusMessage.clear();

    if (! fileSystem)
    {
        outStatusMessage = L"ViewerImgRaw: Active filesystem is missing.";
        return E_FAIL;
    }

    wil::com_ptr<IFileSystemIO> fileIo;
    const HRESULT fileIoHr = fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void());
    if (FAILED(fileIoHr) || ! fileIo)
    {
        const HRESULT hr = FAILED(fileIoHr) ? fileIoHr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        outStatusMessage = std::format(L"ViewerImgRaw: Active filesystem does not implement IFileSystemIO (hr=0x{:08X}).", static_cast<unsigned long>(hr));
        return hr;
    }

    wil::com_ptr<IFileReader> reader;
    const HRESULT openHr = fileIo->CreateFileReader(path.c_str(), reader.put());
    if (FAILED(openHr) || ! reader)
    {
        const HRESULT hr = FAILED(openHr) ? openHr : E_FAIL;
        outStatusMessage = std::format(L"ViewerImgRaw: Failed to create file reader (hr=0x{:08X}).", static_cast<unsigned long>(hr));
        return hr;
    }

    unsigned __int64 sizeBytes = 0;
    const HRESULT sizeHr       = reader->GetSize(&sizeBytes);
    if (FAILED(sizeHr))
    {
        outStatusMessage = std::format(L"ViewerImgRaw: GetSize failed (hr=0x{:08X}).", static_cast<unsigned long>(sizeHr));
        return sizeHr;
    }

    static constexpr uint64_t kMaxRawFileBytes = 1024ull * 1024ull * 1024ull; // 1 GiB
    if (sizeBytes == 0)
    {
        outStatusMessage = L"ViewerImgRaw: File is empty.";
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (sizeBytes > kMaxRawFileBytes)
    {
        outStatusMessage = std::format(L"ViewerImgRaw: File too large ({:L} bytes).", static_cast<uint64_t>(sizeBytes));
        return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }
    if (sizeBytes > static_cast<unsigned __int64>(std::numeric_limits<size_t>::max()))
    {
        outStatusMessage = L"ViewerImgRaw: File too large for address space.";
        return E_OUTOFMEMORY;
    }

    outBytes.resize(static_cast<size_t>(sizeBytes));

    unsigned __int64 pos = 0;
    const HRESULT seekHr = reader->Seek(0, FILE_BEGIN, &pos);
    if (FAILED(seekHr))
    {
        outBytes.clear();
        outStatusMessage = std::format(L"ViewerImgRaw: Seek(FILE_BEGIN, 0) failed (hr=0x{:08X}).", static_cast<unsigned long>(seekHr));
        return seekHr;
    }

    size_t offset = 0;
    while (offset < outBytes.size())
    {
        const unsigned long want = static_cast<unsigned long>(std::min<size_t>(outBytes.size() - offset, 1024u * 1024u));
        unsigned long got        = 0;
        const HRESULT readHr     = reader->Read(outBytes.data() + offset, want, &got);
        if (FAILED(readHr))
        {
            outBytes.clear();
            outStatusMessage = std::format(L"ViewerImgRaw: Read failed (hr=0x{:08X}).", static_cast<unsigned long>(readHr));
            return readHr;
        }
        if (got == 0)
        {
            outBytes.clear();
            outStatusMessage = L"ViewerImgRaw: Unexpected end of file.";
            return HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
        }
        offset += got;
    }

    return S_OK;
}

HRESULT DecodeImageToBgraWic(const uint8_t* data, size_t sizeBytes, uint32_t& outWidth, uint32_t& outHeight, std::vector<uint8_t>& outBgra) noexcept
{
    outWidth  = 0;
    outHeight = 0;
    outBgra.clear();

    if (! data || sizeBytes == 0 || sizeBytes > static_cast<size_t>(std::numeric_limits<DWORD>::max()))
    {
        return E_INVALIDARG;
    }

    wil::com_ptr<IWICImagingFactory> factory;
    HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory2, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(factory.addressof()));
    if (FAILED(hr))
    {
        hr = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(factory.addressof()));
    }
    if (FAILED(hr) || ! factory)
    {
        return FAILED(hr) ? hr : E_FAIL;
    }

    wil::com_ptr<IWICStream> stream;
    hr = factory->CreateStream(stream.addressof());
    if (FAILED(hr) || ! stream)
    {
        return FAILED(hr) ? hr : E_FAIL;
    }

    hr = stream->InitializeFromMemory(const_cast<BYTE*>(reinterpret_cast<const BYTE*>(data)), static_cast<DWORD>(sizeBytes));
    if (FAILED(hr))
    {
        return hr;
    }

    wil::com_ptr<IWICBitmapDecoder> decoder;
    hr = factory->CreateDecoderFromStream(stream.get(), nullptr, WICDecodeMetadataCacheOnLoad, decoder.addressof());
    if (FAILED(hr) || ! decoder)
    {
        return FAILED(hr) ? hr : E_FAIL;
    }

    wil::com_ptr<IWICBitmapFrameDecode> frame;
    hr = decoder->GetFrame(0, frame.addressof());
    if (FAILED(hr) || ! frame)
    {
        return FAILED(hr) ? hr : E_FAIL;
    }

    UINT w = 0;
    UINT h = 0;
    hr     = frame->GetSize(&w, &h);
    if (FAILED(hr) || w == 0 || h == 0)
    {
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    if (w > 16384u || h > 16384u)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    wil::com_ptr<IWICFormatConverter> converter;
    hr = factory->CreateFormatConverter(converter.addressof());
    if (FAILED(hr) || ! converter)
    {
        return FAILED(hr) ? hr : E_FAIL;
    }

    hr = converter->Initialize(frame.get(), GUID_WICPixelFormat32bppBGRA, WICBitmapDitherTypeNone, nullptr, 0.0, WICBitmapPaletteTypeCustom);
    if (FAILED(hr))
    {
        return hr;
    }

    const uint64_t pixelCount = static_cast<uint64_t>(w) * static_cast<uint64_t>(h);
    if (pixelCount == 0 || pixelCount > (static_cast<uint64_t>(std::numeric_limits<size_t>::max()) / 4ull))
    {
        return E_OUTOFMEMORY;
    }

    const UINT stride           = w * 4u;
    const uint64_t bufferSize64 = pixelCount * 4ull;
    if (bufferSize64 > static_cast<uint64_t>(std::numeric_limits<UINT>::max()))
    {
        return E_OUTOFMEMORY;
    }

    const UINT bufferSize = static_cast<UINT>(bufferSize64);
    outBgra.resize(static_cast<size_t>(bufferSize));
    hr = converter->CopyPixels(nullptr, stride, bufferSize, outBgra.data());
    if (FAILED(hr))
    {
        outBgra.clear();
        return hr;
    }

    outWidth  = static_cast<uint32_t>(w);
    outHeight = static_cast<uint32_t>(h);
    return S_OK;
}

struct TurboJpegHeader final
{
    int width  = 0;
    int height = 0;
};

[[nodiscard]] bool TryReadJpegHeaderTurboJpeg(const uint8_t* data, size_t sizeBytes, TurboJpegHeader& out) noexcept
{
    out = {};

    if (! data || sizeBytes == 0 || sizeBytes > static_cast<size_t>(std::numeric_limits<unsigned long>::max()))
    {
        return false;
    }

    wil::unique_any<tjhandle, decltype(&tjDestroy), tjDestroy> handle(tjInitDecompress());
    if (! handle)
    {
        return false;
    }

    int w          = 0;
    int h          = 0;
    int subsamp    = 0;
    int colorspace = 0;
    const int headerRc =
        tjDecompressHeader3(handle.get(), reinterpret_cast<const unsigned char*>(data), static_cast<unsigned long>(sizeBytes), &w, &h, &subsamp, &colorspace);
    if (headerRc != 0 || w <= 0 || h <= 0)
    {
        return false;
    }

    out.width  = w;
    out.height = h;
    return true;
}

[[nodiscard]] bool IsLikelyProgressiveJpeg(const uint8_t* data, size_t sizeBytes) noexcept
{
    if (! data || sizeBytes < 4)
    {
        return false;
    }

    if (data[0] != 0xFF || data[1] != 0xD8)
    {
        return false;
    }

    size_t pos = 2;
    while (pos + 1 < sizeBytes)
    {
        if (data[pos] != 0xFF)
        {
            ++pos;
            continue;
        }

        while (pos < sizeBytes && data[pos] == 0xFF)
        {
            ++pos;
        }
        if (pos >= sizeBytes)
        {
            break;
        }

        const uint8_t marker = data[pos++];
        if (marker == 0xD9 || marker == 0xDA) // EOI / SOS
        {
            break;
        }

        if (marker == 0x01 || (marker >= 0xD0 && marker <= 0xD7)) // TEM / RSTn
        {
            continue;
        }

        if (pos + 1 >= sizeBytes)
        {
            break;
        }

        const uint16_t len = static_cast<uint16_t>(static_cast<uint16_t>(data[pos]) << 8u | static_cast<uint16_t>(data[pos + 1]));
        pos += 2;
        if (len < 2)
        {
            break;
        }

        if (marker == 0xC2) // SOF2 (progressive DCT)
        {
            return true;
        }

        const size_t skip = static_cast<size_t>(len - 2u);
        if (skip > sizeBytes - pos)
        {
            break;
        }
        pos += skip;
    }

    return false;
}

[[nodiscard]] bool ShouldRenderJpegProgressively(const uint8_t* data, size_t sizeBytes, int previewMaxDim) noexcept
{
    TurboJpegHeader header{};
    if (! TryReadJpegHeaderTurboJpeg(data, sizeBytes, header))
    {
        return false;
    }

    if (header.width <= previewMaxDim && header.height <= previewMaxDim)
    {
        return false;
    }

    const bool isProgressive = IsLikelyProgressiveJpeg(data, sizeBytes);
    if (isProgressive)
    {
        return true;
    }

    const uint64_t pixels = static_cast<uint64_t>(header.width) * static_cast<uint64_t>(header.height);
    if (pixels >= 12'000'000ull)
    {
        return true;
    }

    return sizeBytes >= 2ull * 1024ull * 1024ull;
}

struct TurboJpegScaledDims final
{
    int width  = 0;
    int height = 0;
};

[[nodiscard]] TurboJpegScaledDims ChooseTurboJpegScaledDims(int width, int height, int maxDim) noexcept
{
    TurboJpegScaledDims out{};
    out.width  = width;
    out.height = height;

    if (width <= 0 || height <= 0 || maxDim <= 0)
    {
        return out;
    }

    if (width <= maxDim && height <= maxDim)
    {
        return out;
    }

    int factorCount                = 0;
    const tjscalingfactor* factors = tjGetScalingFactors(&factorCount);
    if (! factors || factorCount <= 0)
    {
        return out;
    }

    long long bestPixels = -1;
    for (int i = 0; i < factorCount; ++i)
    {
        const int scaledW = TJSCALED(width, factors[i]);
        const int scaledH = TJSCALED(height, factors[i]);
        if (scaledW <= 0 || scaledH <= 0)
        {
            continue;
        }
        if (scaledW > maxDim || scaledH > maxDim)
        {
            continue;
        }

        const long long pixels = static_cast<long long>(scaledW) * static_cast<long long>(scaledH);
        if (pixels > bestPixels)
        {
            bestPixels = pixels;
            out.width  = scaledW;
            out.height = scaledH;
        }
    }

    if (bestPixels >= 0)
    {
        return out;
    }

    long long bestMaxSide = std::numeric_limits<long long>::max();
    for (int i = 0; i < factorCount; ++i)
    {
        const int scaledW = TJSCALED(width, factors[i]);
        const int scaledH = TJSCALED(height, factors[i]);
        if (scaledW <= 0 || scaledH <= 0)
        {
            continue;
        }
        const long long maxSide = static_cast<long long>(std::max(scaledW, scaledH));
        if (maxSide < bestMaxSide)
        {
            bestMaxSide = maxSide;
            out.width   = scaledW;
            out.height  = scaledH;
        }
    }

    return out;
}

HRESULT DecodeJpegToBgraTurboJpegScaled(
    const uint8_t* data, size_t sizeBytes, int maxDim, uint32_t& outWidth, uint32_t& outHeight, std::vector<uint8_t>& outBgra) noexcept
{
    outWidth  = 0;
    outHeight = 0;
    outBgra.clear();

    if (! data || sizeBytes == 0 || sizeBytes > static_cast<size_t>(std::numeric_limits<unsigned long>::max()))
    {
        return E_INVALIDARG;
    }

    wil::unique_any<tjhandle, decltype(&tjDestroy), tjDestroy> handle(tjInitDecompress());
    if (! handle)
    {
        return E_FAIL;
    }

    int w          = 0;
    int h          = 0;
    int subsamp    = 0;
    int colorspace = 0;
    const int headerRc =
        tjDecompressHeader3(handle.get(), reinterpret_cast<const unsigned char*>(data), static_cast<unsigned long>(sizeBytes), &w, &h, &subsamp, &colorspace);
    if (headerRc != 0 || w <= 0 || h <= 0)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const TurboJpegScaledDims scaled = ChooseTurboJpegScaledDims(w, h, maxDim);
    if (scaled.width <= 0 || scaled.height <= 0)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (scaled.width > maxDim || scaled.height > maxDim)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const uint64_t pixelCount = static_cast<uint64_t>(scaled.width) * static_cast<uint64_t>(scaled.height);
    if (pixelCount == 0 || pixelCount > (static_cast<uint64_t>(std::numeric_limits<size_t>::max()) / 4ull))
    {
        return E_OUTOFMEMORY;
    }

    outBgra.resize(static_cast<size_t>(pixelCount) * 4u);

    constexpr int flags = TJFLAG_FASTDCT | TJFLAG_FASTUPSAMPLE;
    const int rc        = tjDecompress2(handle.get(),
                                 reinterpret_cast<const unsigned char*>(data),
                                 static_cast<unsigned long>(sizeBytes),
                                 outBgra.data(),
                                 scaled.width,
                                 0,
                                 scaled.height,
                                 TJPF_BGRA,
                                 flags);
    if (rc != 0)
    {
        outBgra.clear();
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    outWidth  = static_cast<uint32_t>(scaled.width);
    outHeight = static_cast<uint32_t>(scaled.height);
    return S_OK;
}

HRESULT DecodeJpegToBgraTurboJpeg(const uint8_t* data, size_t sizeBytes, uint32_t& outWidth, uint32_t& outHeight, std::vector<uint8_t>& outBgra) noexcept
{
    static constexpr int kMaxJpegDim = 16384;
    return DecodeJpegToBgraTurboJpegScaled(data, sizeBytes, kMaxJpegDim, outWidth, outHeight, outBgra);
}

bool TryDecodeRawEmbeddedThumbnailFromBufferToBgra(const std::vector<uint8_t>& fileBytes,
                                                   uint32_t& outWidth,
                                                   uint32_t& outHeight,
                                                   std::vector<uint8_t>& outBgra,
                                                   bool& outThumbAvailable,
                                                   ExifData& outExif) noexcept
{
    outWidth  = 0;
    outHeight = 0;
    outBgra.clear();
    outThumbAvailable = false;
    outExif           = {};

    if (fileBytes.empty())
    {
        return false;
    }

    LibRaw raw;
    const int openRet = raw.open_buffer(fileBytes.data(), fileBytes.size());
    if (openRet != LIBRAW_SUCCESS)
    {
        return false;
    }
    auto recycle = wil::scope_exit([&] { raw.recycle(); });

    const int thumbRet = raw.unpack_thumb();
    if (thumbRet != LIBRAW_SUCCESS)
    {
        return false;
    }

    outThumbAvailable = true;
    outExif           = ExtractExifData(raw);

    const libraw_thumbnail_t& thumb = raw.imgdata.thumbnail;
    if (thumb.twidth == 0 || thumb.theight == 0 || thumb.tlength == 0 || thumb.thumb == nullptr)
    {
        return false;
    }

    if (thumb.tformat == LIBRAW_THUMBNAIL_JPEG)
    {
        const ExifData jpegExif = ExtractExifFromJpeg(reinterpret_cast<const uint8_t*>(thumb.thumb), static_cast<size_t>(thumb.tlength));
        outExif.orientation     = jpegExif.orientation;
        if (outExif.camera.empty() && ! jpegExif.camera.empty())
        {
            outExif.camera = jpegExif.camera;
        }
        if (outExif.lens.empty() && ! jpegExif.lens.empty())
        {
            outExif.lens = jpegExif.lens;
        }
        if (outExif.dateTime.empty() && ! jpegExif.dateTime.empty())
        {
            outExif.dateTime = jpegExif.dateTime;
        }
        if (outExif.iso <= 0.0f && jpegExif.iso > 0.0f)
        {
            outExif.iso = jpegExif.iso;
        }
        if (outExif.shutterSeconds <= 0.0f && jpegExif.shutterSeconds > 0.0f)
        {
            outExif.shutterSeconds = jpegExif.shutterSeconds;
        }
        if (outExif.aperture <= 0.0f && jpegExif.aperture > 0.0f)
        {
            outExif.aperture = jpegExif.aperture;
        }
        if (outExif.focalLengthMm <= 0.0f && jpegExif.focalLengthMm > 0.0f)
        {
            outExif.focalLengthMm = jpegExif.focalLengthMm;
        }
        outExif.valid = outExif.valid || jpegExif.valid || outExif.orientation != 1;

        const HRESULT hr =
            DecodeJpegToBgraTurboJpeg(reinterpret_cast<const uint8_t*>(thumb.thumb), static_cast<size_t>(thumb.tlength), outWidth, outHeight, outBgra);
        return SUCCEEDED(hr);
    }

    if (thumb.tformat != LIBRAW_THUMBNAIL_BITMAP && thumb.tformat != LIBRAW_THUMBNAIL_BITMAP16)
    {
        return false;
    }

    const uint32_t w      = static_cast<uint32_t>(thumb.twidth);
    const uint32_t h      = static_cast<uint32_t>(thumb.theight);
    const uint32_t colors = static_cast<uint32_t>(std::max(thumb.tcolors, 0));
    const uint32_t bits   = (thumb.tformat == LIBRAW_THUMBNAIL_BITMAP16) ? 16u : 8u;

    if (w == 0 || h == 0 || colors == 0)
    {
        return false;
    }

    const uint64_t pixelCount = static_cast<uint64_t>(w) * static_cast<uint64_t>(h);
    if (pixelCount == 0 || pixelCount > (static_cast<uint64_t>(std::numeric_limits<size_t>::max()) / 4ull))
    {
        return false;
    }

    const size_t bytesPerSample    = bits == 16 ? 2u : 1u;
    const size_t expectedDataBytes = static_cast<size_t>(pixelCount) * static_cast<size_t>(colors) * bytesPerSample;
    if (static_cast<size_t>(thumb.tlength) < expectedDataBytes)
    {
        return false;
    }

    outBgra.resize(static_cast<size_t>(pixelCount) * 4u);

    if (bits == 8)
    {
        const uint8_t* src = reinterpret_cast<const uint8_t*>(thumb.thumb);
        for (uint64_t i = 0; i < pixelCount; ++i)
        {
            const size_t si = static_cast<size_t>(i) * colors;
            const uint8_t r = colors >= 1 ? src[si + 0] : uint8_t{0};
            const uint8_t g = colors >= 2 ? src[si + 1] : r;
            const uint8_t b = colors >= 3 ? src[si + 2] : r;

            const size_t di = static_cast<size_t>(i) * 4u;
            outBgra[di + 0] = b;
            outBgra[di + 1] = g;
            outBgra[di + 2] = r;
            outBgra[di + 3] = uint8_t{255};
        }
    }
    else
    {
        const uint16_t* src = reinterpret_cast<const uint16_t*>(thumb.thumb);
        for (uint64_t i = 0; i < pixelCount; ++i)
        {
            const size_t si    = static_cast<size_t>(i) * colors;
            const uint16_t r16 = colors >= 1 ? src[si + 0] : uint16_t{0};
            const uint16_t g16 = colors >= 2 ? src[si + 1] : r16;
            const uint16_t b16 = colors >= 3 ? src[si + 2] : r16;

            const uint8_t r = static_cast<uint8_t>(r16 >> 8);
            const uint8_t g = static_cast<uint8_t>(g16 >> 8);
            const uint8_t b = static_cast<uint8_t>(b16 >> 8);

            const size_t di = static_cast<size_t>(i) * 4u;
            outBgra[di + 0] = b;
            outBgra[di + 1] = g;
            outBgra[di + 2] = r;
            outBgra[di + 3] = uint8_t{255};
        }
    }

    outWidth  = w;
    outHeight = h;
    return true;
}

HRESULT DecodeRawFullImageFromBufferToBgra(const RawDecodeSettings& cfg,
                                           const std::vector<uint8_t>& fileBytes,
                                           uint32_t& outWidth,
                                           uint32_t& outHeight,
                                           std::vector<uint8_t>& outBgra,
                                           std::wstring& outStatusMessage,
                                           ExifData& outExif,
                                           const LibRawProgressHost* progressHost) noexcept
{
    outWidth  = 0;
    outHeight = 0;
    outBgra.clear();
    outStatusMessage.clear();
    outExif = {};

    if (fileBytes.empty())
    {
        outStatusMessage = L"ViewerImgRaw: File is empty.";
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    LibRaw raw;
    raw.imgdata.params.half_size      = cfg.halfSize ? 1 : 0;
    raw.imgdata.params.use_camera_wb  = cfg.useCameraWb ? 1 : 0;
    raw.imgdata.params.use_auto_wb    = cfg.autoWb ? 1 : 0;
    raw.imgdata.params.no_auto_bright = 1;
    raw.imgdata.params.output_bps     = 8;

    LibRawProgressContext progressCtx{};
    if (progressHost && progressHost->requestIdCounter && progressHost->hwnd)
    {
        progressCtx.host = *progressHost;
        raw.set_progress_handler(&LibRawProgressCallback, &progressCtx);
    }

    const int openRet = raw.open_buffer(fileBytes.data(), fileBytes.size());
    if (openRet != LIBRAW_SUCCESS)
    {
        outStatusMessage = std::format(L"ViewerImgRaw: LibRaw open failed: {} (code={}).", WideFromUtf8(LibRaw::strerror(openRet)), openRet);
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const int unpackRet = raw.unpack();
    if (unpackRet != LIBRAW_SUCCESS)
    {
        outStatusMessage = std::format(L"ViewerImgRaw: LibRaw unpack failed: {} (code={}).", WideFromUtf8(LibRaw::strerror(unpackRet)), unpackRet);
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    outExif = ExtractExifData(raw);

    const int processRet = raw.dcraw_process();
    if (processRet != LIBRAW_SUCCESS)
    {
        outStatusMessage = std::format(L"ViewerImgRaw: LibRaw process failed: {} (code={}).", WideFromUtf8(LibRaw::strerror(processRet)), processRet);
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    int memErr                       = 0;
    libraw_processed_image_t* memImg = raw.dcraw_make_mem_image(&memErr);
    if (! memImg || memErr != LIBRAW_SUCCESS)
    {
        outStatusMessage = std::format(L"ViewerImgRaw: LibRaw make_mem_image failed: {} (code={}).", WideFromUtf8(LibRaw::strerror(memErr)), memErr);
        if (memImg)
        {
            LibRaw::dcraw_clear_mem(memImg);
        }
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    auto freeMemImg = wil::scope_exit(
        [&]
        {
            LibRaw::dcraw_clear_mem(memImg);
            raw.recycle();
        });

    const uint32_t w      = static_cast<uint32_t>(memImg->width);
    const uint32_t h      = static_cast<uint32_t>(memImg->height);
    const uint32_t colors = static_cast<uint32_t>(memImg->colors);
    const uint32_t bits   = static_cast<uint32_t>(memImg->bits);

    if (w == 0 || h == 0 || colors == 0)
    {
        outStatusMessage = L"ViewerImgRaw: Invalid decoded image dimensions.";
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    static constexpr uint32_t kMaxBitmapDim = 16384u;
    if (w > kMaxBitmapDim || h > kMaxBitmapDim)
    {
        outStatusMessage = std::format(L"ViewerImgRaw: Image too large ({}×{}).", w, h);
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const uint64_t pixelCount = static_cast<uint64_t>(w) * static_cast<uint64_t>(h);
    if (pixelCount == 0 || pixelCount > (static_cast<uint64_t>(std::numeric_limits<size_t>::max()) / 4ull))
    {
        outStatusMessage = L"ViewerImgRaw: Decoded image is too large.";
        return E_OUTOFMEMORY;
    }

    const size_t outBytes = static_cast<size_t>(pixelCount) * 4u;
    outBgra.assign(outBytes, 0);

    const size_t bytesPerSample    = bits == 16 ? 2u : 1u;
    const size_t expectedDataBytes = static_cast<size_t>(pixelCount) * static_cast<size_t>(colors) * bytesPerSample;
    if (memImg->data_size < expectedDataBytes)
    {
        outStatusMessage = L"ViewerImgRaw: Decoded image buffer is truncated.";
        outBgra.clear();
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    if (bits == 8)
    {
        const uint8_t* src = memImg->data;
        for (uint64_t i = 0; i < pixelCount; ++i)
        {
            const size_t si = static_cast<size_t>(i) * colors;
            const uint8_t r = colors >= 1 ? src[si + 0] : uint8_t{0};
            const uint8_t g = colors >= 2 ? src[si + 1] : r;
            const uint8_t b = colors >= 3 ? src[si + 2] : r;

            const size_t di = static_cast<size_t>(i) * 4u;
            outBgra[di + 0] = b;
            outBgra[di + 1] = g;
            outBgra[di + 2] = r;
            outBgra[di + 3] = 255u;
        }
    }
    else if (bits == 16)
    {
        const uint16_t* src = reinterpret_cast<const uint16_t*>(memImg->data);
        for (uint64_t i = 0; i < pixelCount; ++i)
        {
            const size_t si    = static_cast<size_t>(i) * colors;
            const uint16_t r16 = colors >= 1 ? src[si + 0] : uint16_t{0};
            const uint16_t g16 = colors >= 2 ? src[si + 1] : r16;
            const uint16_t b16 = colors >= 3 ? src[si + 2] : r16;

            const uint8_t r = static_cast<uint8_t>(r16 >> 8);
            const uint8_t g = static_cast<uint8_t>(g16 >> 8);
            const uint8_t b = static_cast<uint8_t>(b16 >> 8);

            const size_t di = static_cast<size_t>(i) * 4u;
            outBgra[di + 0] = b;
            outBgra[di + 1] = g;
            outBgra[di + 2] = r;
            outBgra[di + 3] = 255u;
        }
    }
    else
    {
        outStatusMessage = std::format(L"ViewerImgRaw: Unsupported bit depth ({}).", bits);
        outBgra.clear();
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    outWidth  = w;
    outHeight = h;
    return S_OK;
}
} // namespace

void ViewerImgRaw::OnAsyncProgress(int stage, int percent) noexcept
{
    _rawProgressStage     = stage;
    _rawProgressPercent   = percent;
    _rawProgressStageText = WideFromUtf8(LibRaw::strprogress(static_cast<LibRaw_progress>(stage)));

    const HWND hwnd = _hWnd.get();
    if (hwnd)
    {
        InvalidateRect(hwnd, &_statusRect, FALSE);
        InvalidateRect(hwnd, &_contentRect, FALSE);
    }
}

void ViewerImgRaw::ClearImageCache() noexcept
{
    static_cast<void>(_openRequestId.fetch_add(1, std::memory_order_acq_rel));
    EndLoadingUi();

    {
        std::scoped_lock lock(_cacheMutex);
        _inflightDecodes.clear();
        _imageCache.clear();
    }

    _currentImageOwned.reset();
    _currentImage = nullptr;
    _currentImageKey.clear();
    _imageBitmap.reset();
    _exifOverlayText.clear();
    _rawProgressPercent = -1;
    _rawProgressStage   = -1;
    _rawProgressStageText.clear();
    _panOffsetXPx            = 0.0f;
    _panOffsetYPx            = 0.0f;
    _panning                 = false;
    _baseOrientation         = 1;
    _userOrientation         = 1;
    _viewOrientation         = 1;
    _orientationUserModified = false;
}

bool ViewerImgRaw::TryUseCachedImage(HWND hwnd, const std::wstring& path, bool& outContinueDecoding) noexcept
{
    outContinueDecoding = false;
    if (! hwnd || path.empty())
    {
        return false;
    }

    CachedImage* cached = nullptr;
    {
        std::scoped_lock lock(_cacheMutex);
        const auto it = _imageCache.find(path);
        if (it == _imageCache.end() || ! it->second)
        {
            return false;
        }
        cached = it->second.get();
    }

    if (! cached)
    {
        return false;
    }

    const bool hasRaw   = cached->rawWidth != 0 && cached->rawHeight != 0 && ! cached->rawBgra.empty();
    const bool hasThumb = cached->thumbDecoded && cached->thumbWidth != 0 && cached->thumbHeight != 0 && ! cached->thumbBgra.empty();

    if (_displayMode == DisplayMode::Raw)
    {
        if (hasRaw)
        {
            _displayedMode      = DisplayMode::Raw;
            outContinueDecoding = false;
        }
        else if (hasThumb)
        {
            _displayedMode      = DisplayMode::Thumbnail;
            outContinueDecoding = true;
        }
        else
        {
            return false;
        }
    }
    else
    {
        if (hasThumb)
        {
            _displayedMode      = DisplayMode::Thumbnail;
            outContinueDecoding = false;
        }
        else if (hasRaw)
        {
            _displayedMode      = DisplayMode::Raw;
            outContinueDecoding = cached->thumbAvailable || ! _currentSidecarJpegPath.empty();
        }
        else
        {
            return false;
        }
    }

    _statusMessage.clear();
    _imageBitmap.reset();

    _currentImageOwned.reset();
    _currentImage    = cached;
    _currentImageKey = path;
    UpdateOrientationState();
    RebuildExifOverlayText();

    if (_hostAlerts)
    {
        static_cast<void>(_hostAlerts->ClearAlert(HOST_ALERT_SCOPE_WINDOW, nullptr));
    }
    _alertVisible = false;

    return true;
}

bool ViewerImgRaw::HasDisplayImage() const noexcept
{
    const CachedImage* image = _currentImage;
    if (! image)
    {
        return false;
    }

    if (IsDisplayingThumbnail())
    {
        return image->thumbDecoded && image->thumbWidth != 0 && image->thumbHeight != 0 && ! image->thumbBgra.empty();
    }

    return image->rawWidth != 0 && image->rawHeight != 0 && ! image->rawBgra.empty();
}

bool ViewerImgRaw::IsDisplayingThumbnail() const noexcept
{
    return _displayedMode == DisplayMode::Thumbnail;
}

void ViewerImgRaw::SetDisplayMode(DisplayMode mode) noexcept
{
    if (_displayMode == mode)
    {
        return;
    }

    _displayMode = mode;

    const HWND hwnd = _hWnd.get();
    if (! hwnd || _currentPath.empty())
    {
        return;
    }

    const CachedImage* image = _currentImage;
    if (image && _currentImageKey == _currentPath)
    {
        const bool hasRaw   = image->rawWidth != 0 && image->rawHeight != 0 && ! image->rawBgra.empty();
        const bool hasThumb = image->thumbDecoded && image->thumbWidth != 0 && image->thumbHeight != 0 && ! image->thumbBgra.empty();

        if (_displayMode == DisplayMode::Raw && hasRaw)
        {
            _displayedMode = DisplayMode::Raw;
            _imageBitmap.reset();
            UpdateOrientationState();
            RebuildExifOverlayText();
            _panOffsetXPx = 0.0f;
            _panOffsetYPx = 0.0f;
            _panning      = false;
            UpdateScrollBars(hwnd);
            return;
        }

        if (_displayMode == DisplayMode::Thumbnail && hasThumb)
        {
            _displayedMode = DisplayMode::Thumbnail;
            _imageBitmap.reset();
            UpdateOrientationState();
            RebuildExifOverlayText();
            _panOffsetXPx = 0.0f;
            _panOffsetYPx = 0.0f;
            _panning      = false;
            UpdateScrollBars(hwnd);
            return;
        }
    }

    StartAsyncOpen(hwnd, _currentPath, false);
}

void ViewerImgRaw::UpdateOrientationState() noexcept
{
    const CachedImage* image = _currentImage;
    uint16_t base            = 1;
    if (image)
    {
        base = IsDisplayingThumbnail() ? image->thumbOrientation : image->rawOrientation;
    }

    _baseOrientation         = NormalizeExifOrientation(base);
    _userOrientation         = NormalizeExifOrientation(_userOrientation);
    _viewOrientation         = ComposeExifOrientation(_userOrientation, _baseOrientation);
    _orientationUserModified = (_userOrientation != 1);
}

void ViewerImgRaw::RebuildExifOverlayText() noexcept
{
    _exifOverlayText.clear();

    if (! _showExifOverlay || ! HasDisplayImage())
    {
        return;
    }

    const CachedImage* image = _currentImage;
    if (! image || ! image->exif.valid)
    {
        return;
    }

    const uint32_t w           = IsDisplayingThumbnail() ? image->thumbWidth : image->rawWidth;
    const uint32_t h           = IsDisplayingThumbnail() ? image->thumbHeight : image->rawHeight;
    const uint16_t orientation = (_viewOrientation >= 1 && _viewOrientation <= 8) ? _viewOrientation : static_cast<uint16_t>(1);
    const bool swapAxes        = (orientation >= 5 && orientation <= 8);
    const uint32_t dispW       = swapAxes ? h : w;
    const uint32_t dispH       = swapAxes ? w : h;

    std::wstring src;
    if (IsDisplayingThumbnail())
    {
        switch (image->thumbSource)
        {
            case CachedImage::ThumbSource::SidecarJpeg: src = L"JPG"; break;
            case CachedImage::ThumbSource::Embedded: src = L"THUMB"; break;
            case CachedImage::ThumbSource::None:
            default: src = L"THUMB"; break;
        }
    }
    else
    {
        const std::wstring extLower = ToLowerCopy(PathExtensionView(_currentPath));
        if (IsJpegExtension(extLower))
        {
            src = L"JPG";
        }
        else if (_otherIndex < _otherItems.size() && _otherItems[_otherIndex].isRaw)
        {
            src = L"RAW";
        }
        else if (! extLower.empty() && extLower.size() <= 8)
        {
            src.assign(extLower);
            for (wchar_t& ch : src)
            {
                ch = static_cast<wchar_t>(std::towupper(ch));
            }
            if (! src.empty() && src.front() == L'.')
            {
                src.erase(src.begin());
            }
        }
        else
        {
            src = L"IMG";
        }
    }

    std::wstring line1;
    if (! image->exif.camera.empty())
    {
        line1 = image->exif.camera;
    }

    std::wstring line2;
    if (! image->exif.lens.empty())
    {
        line2 = image->exif.lens;
    }

    std::wstring line3;
    std::wstring details;
    if (image->exif.iso > 0.0f)
    {
        details.append(std::format(L"ISO {:g}", image->exif.iso));
    }
    if (image->exif.shutterSeconds > 0.0f)
    {
        if (! details.empty())
        {
            details.append(L"  ");
        }
        details.append(std::format(L"{:g}s", image->exif.shutterSeconds));
    }
    if (image->exif.aperture > 0.0f)
    {
        if (! details.empty())
        {
            details.append(L"  ");
        }
        details.append(std::format(L"f/{:g}", image->exif.aperture));
    }
    if (image->exif.focalLengthMm > 0.0f)
    {
        if (! details.empty())
        {
            details.append(L"  ");
        }
        details.append(std::format(L"{:g}mm", image->exif.focalLengthMm));
    }
    line3 = std::move(details);

    std::wstring line4;
    if (! image->exif.dateTime.empty())
    {
        line4 = image->exif.dateTime;
    }

    std::wstring text;
    text.reserve(256);

    if (! line1.empty())
    {
        text.append(line1);
        text.push_back(L'\n');
    }
    if (! line2.empty())
    {
        text.append(line2);
        text.push_back(L'\n');
    }
    if (! line3.empty())
    {
        text.append(line3);
        text.push_back(L'\n');
    }
    if (! line4.empty())
    {
        text.append(line4);
        text.push_back(L'\n');
    }

    text.append(std::format(L"{}  {}×{}", src.empty() ? L"IMG" : src, dispW, dispH));
    _exifOverlayText = std::move(text);
}

std::wstring ViewerImgRaw::BuildStatusBarText(bool drewImage, float displayedZoom) const
{
    static_cast<void>(displayedZoom);

    if (_isLoading)
    {
        std::wstring text = LoadStringResource(g_hInstance, IDS_VIEWERRAW_STATUS_LOADING);
        if (_rawProgressStage >= 0)
        {
            if (! text.empty())
            {
                text.append(L"  ");
            }
            text.append(_rawProgressStageText.empty() ? L"RAW" : _rawProgressStageText);

            if (_rawProgressPercent >= 0)
            {
                text.append(std::format(L" {}%", _rawProgressPercent));
            }
        }

        return text;
    }

    if (! drewImage || ! HasDisplayImage() || ! _currentImage)
    {
        return _statusMessage.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERRAW_STATUS_NO_IMAGE) : _statusMessage;
    }

    std::wstring pos;
    if (! _otherItems.empty())
    {
        pos = std::format(L"{}/{}", (_otherIndex + 1), _otherItems.size());
    }

    std::wstring out;
    out.reserve(256);

    if (! pos.empty())
    {
        out.append(pos);
        out.append(L"  ");
    }

    if (! _currentLabel.empty())
    {
        out.append(_currentLabel);
        out.append(L"  ");
    }

    if (! _statusMessage.empty())
    {
        out.append(L"  ");
        out.append(_statusMessage);
    }

    return out;
}

std::wstring ViewerImgRaw::BuildStatusBarRightText(bool drewImage, float displayedZoom) const
{
    if (! drewImage || ! HasDisplayImage() || ! _currentImage)
    {
        return {};
    }

    const CachedImage* image = _currentImage;
    const bool thumb         = IsDisplayingThumbnail();
    const uint32_t w         = thumb ? image->thumbWidth : image->rawWidth;
    const uint32_t h         = thumb ? image->thumbHeight : image->rawHeight;
    if (w == 0 || h == 0)
    {
        return {};
    }

    const uint16_t orientation = (_viewOrientation >= 1 && _viewOrientation <= 8) ? _viewOrientation : static_cast<uint16_t>(1);
    const bool swapAxes        = (orientation >= 5 && orientation <= 8);
    const uint32_t dispW       = swapAxes ? h : w;
    const uint32_t dispH       = swapAxes ? w : h;

    std::wstring src;
    if (thumb)
    {
        switch (image->thumbSource)
        {
            case CachedImage::ThumbSource::SidecarJpeg: src = L"JPG"; break;
            case CachedImage::ThumbSource::Embedded: src = L"THUMB"; break;
            case CachedImage::ThumbSource::None:
            default: src = L"THUMB"; break;
        }
    }
    else
    {
        const std::wstring extLower = ToLowerCopy(PathExtensionView(_currentPath));
        if (IsJpegExtension(extLower))
        {
            src = L"JPG";
        }
        else if (_otherIndex < _otherItems.size() && _otherItems[_otherIndex].isRaw)
        {
            src = L"RAW";
        }
        else if (! extLower.empty() && extLower.size() <= 8)
        {
            src.assign(extLower);
            for (wchar_t& ch : src)
            {
                ch = static_cast<wchar_t>(std::towupper(ch));
            }
            if (! src.empty() && src.front() == L'.')
            {
                src.erase(src.begin());
            }
        }
        else
        {
            src = L"IMG";
        }
    }

    const int zoomPercent = static_cast<int>(std::lround(std::clamp(displayedZoom, 0.01f, 64.0f) * 100.0f));

    std::wstring details = std::format(L"{}  {}×{}  {}%", src.empty() ? L"IMG" : src, dispW, dispH, zoomPercent);

    if (_orientationUserModified)
    {
        details.append(L"  Ori*");
    }

    bool anyAdjust = false;
    if (std::fabs(_brightness) > 0.001f)
    {
        details.append(std::format(L"  B{:+.2f}", _brightness));
        anyAdjust = true;
    }
    if (std::fabs(_contrast - 1.0f) > 0.001f)
    {
        details.append(std::format(L"  C{:.2f}", _contrast));
        anyAdjust = true;
    }
    if (std::fabs(_gamma - 1.0f) > 0.001f)
    {
        details.append(std::format(L"  G{:.2f}", _gamma));
        anyAdjust = true;
    }
    if (_grayscale)
    {
        details.append(L"  Gray");
        anyAdjust = true;
    }
    if (_negative)
    {
        details.append(L"  Neg");
        anyAdjust = true;
    }

    static_cast<void>(anyAdjust);
    return details;
}

void ViewerImgRaw::UpdateNeighborCache(uint64_t requestId) noexcept
{
    if (_config.prevCache == 0 && _config.nextCache == 0)
    {
        return;
    }

    if (_otherItems.size() <= 1)
    {
        return;
    }

    const size_t count = _otherItems.size();
    if (_otherIndex >= count)
    {
        return;
    }

    const size_t reserveCount =
        1ull + static_cast<size_t>(std::min<uint32_t>(_config.prevCache, 8u)) + static_cast<size_t>(std::min<uint32_t>(_config.nextCache, 8u));

    std::unordered_set<std::wstring> keep;
    keep.reserve(reserveCount);
    keep.insert(_otherItems[_otherIndex].primaryPath);

    const uint32_t prevN = std::min<uint32_t>(_config.prevCache, static_cast<uint32_t>(count - 1));
    const uint32_t nextN = std::min<uint32_t>(_config.nextCache, static_cast<uint32_t>(count - 1));

    for (uint32_t i = 1; i <= prevN; ++i)
    {
        const size_t idx = (_otherIndex + count - i) % count;
        keep.insert(_otherItems[idx].primaryPath);
    }

    for (uint32_t i = 1; i <= nextN; ++i)
    {
        const size_t idx = (_otherIndex + i) % count;
        keep.insert(_otherItems[idx].primaryPath);
    }

    std::scoped_lock lock(_cacheMutex);
    for (auto it = _imageCache.begin(); it != _imageCache.end();)
    {
        if (keep.find(it->first) == keep.end())
        {
            it = _imageCache.erase(it);
        }
        else
        {
            ++it;
        }
    }

    StartPrefetchNeighbors(requestId);
}

void ViewerImgRaw::StartPrefetchNeighbors(uint64_t requestId) noexcept
{
    if (_config.prevCache == 0 && _config.nextCache == 0)
    {
        return;
    }

    if (_otherItems.size() <= 1)
    {
        return;
    }

    const size_t count = _otherItems.size();
    if (_otherIndex >= count)
    {
        return;
    }

    const wil::com_ptr<IFileSystem> fileSystem = _fileSystem;
    if (! fileSystem)
    {
        return;
    }

    struct PrefetchItem final
    {
        std::wstring primaryPath;
        std::wstring sidecarJpegPath;
        bool isRaw = false;
    };

    std::vector<PrefetchItem> items;
    const uint32_t prevN = std::min<uint32_t>(_config.prevCache, static_cast<uint32_t>(count - 1));
    const uint32_t nextN = std::min<uint32_t>(_config.nextCache, static_cast<uint32_t>(count - 1));

    items.reserve(static_cast<size_t>(prevN) + static_cast<size_t>(nextN));
    for (uint32_t i = 1; i <= prevN; ++i)
    {
        const size_t idx = (_otherIndex + count - i) % count;
        PrefetchItem prefetch{};
        const OtherItem& other   = _otherItems[idx];
        prefetch.primaryPath     = other.primaryPath;
        prefetch.sidecarJpegPath = other.sidecarJpegPath;
        prefetch.isRaw           = other.isRaw;
        items.push_back(std::move(prefetch));
    }

    for (uint32_t i = 1; i <= nextN; ++i)
    {
        const size_t idx = (_otherIndex + i) % count;
        PrefetchItem prefetch{};
        const OtherItem& other   = _otherItems[idx];
        prefetch.primaryPath     = other.primaryPath;
        prefetch.sidecarJpegPath = other.sidecarJpegPath;
        prefetch.isRaw           = other.isRaw;
        items.push_back(std::move(prefetch));
    }

    const DisplayMode prefetchMode = _displayMode;
    RawDecodeSettings decodeCfg{};
    decodeCfg.halfSize    = _config.halfSize;
    decodeCfg.useCameraWb = _config.useCameraWb;
    decodeCfg.autoWb      = _config.autoWb;

    AddRef();
    struct PrefetchWorkItem final
    {
        PrefetchWorkItem()                                   = default;
        PrefetchWorkItem(const PrefetchWorkItem&)            = delete;
        PrefetchWorkItem& operator=(const PrefetchWorkItem&) = delete;

        wil::unique_hmodule moduleKeepAlive;
        std::function<void()> work;
    };

    auto ctx = std::unique_ptr<PrefetchWorkItem>(new (std::nothrow) PrefetchWorkItem{});

    ctx->moduleKeepAlive = AcquireModuleReferenceFromAddress(&kViewerImgRawModuleAnchor);
    ctx->work            = [this, fileSystem, items = std::move(items), requestId, prefetchMode, decodeCfg]() mutable
    {
        auto releaseSelf = wil::scope_exit([&] { Release(); });

        for (size_t index = 0; index < items.size(); ++index)
        {
            if (requestId != _openRequestId.load(std::memory_order_acquire))
            {
                return;
            }

            const PrefetchItem& item = items[index];
            const std::wstring& path = item.primaryPath;

            if (path.empty())
            {
                continue;
            }

            {
                std::scoped_lock lock(_cacheMutex);
                if (_imageCache.find(path) != _imageCache.end())
                {
                    continue;
                }
                if (_inflightDecodes.find(path) != _inflightDecodes.end())
                {
                    continue;
                }
                _inflightDecodes.insert(path);
            }

            if (requestId != _openRequestId.load(std::memory_order_acquire))
            {
                std::scoped_lock lock(_cacheMutex);
                _inflightDecodes.erase(path);
                return;
            }

            std::vector<uint8_t> fileBytes;
            std::wstring readStatus;
            const HRESULT readHr = ReadFileAllBytes(fileSystem.get(), path, fileBytes, readStatus);
            if (FAILED(readHr))
            {
                std::scoped_lock lock(_cacheMutex);
                _inflightDecodes.erase(path);
                continue;
            }

            bool hasRawAlready = false;
            {
                std::scoped_lock lock(_cacheMutex);
                const auto it = _imageCache.find(path);
                if (it != _imageCache.end() && it->second)
                {
                    const CachedImage* cached = it->second.get();
                    hasRawAlready             = cached->rawWidth != 0 && cached->rawHeight != 0 && ! cached->rawBgra.empty();
                }
            }

            uint32_t rawW = 0;
            uint32_t rawH = 0;
            std::vector<uint8_t> rawBgra;
            std::wstring rawStatus;
            ExifData rawExif{};
            HRESULT rawHr = S_OK;

            bool thumbAvailable = ! item.sidecarJpegPath.empty();
            ExifData thumbExif{};
            uint32_t thumbW = 0;
            uint32_t thumbH = 0;
            std::vector<uint8_t> thumbBgra;
            bool thumbDecoded                    = false;
            CachedImage::ThumbSource thumbSource = CachedImage::ThumbSource::None;

            if (! item.isRaw)
            {
                if (! hasRawAlready)
                {
                    const std::wstring_view ext = PathExtensionView(path);
                    if (IsJpegExtension(ext))
                    {
                        rawHr   = DecodeJpegToBgraTurboJpeg(fileBytes.data(), fileBytes.size(), rawW, rawH, rawBgra);
                        rawExif = ExtractExifFromJpeg(fileBytes.data(), fileBytes.size());
                    }
                    else
                    {
                        rawHr = DecodeImageToBgraWic(fileBytes.data(), fileBytes.size(), rawW, rawH, rawBgra);
                    }
                }
            }
            else
            {
                if (! item.sidecarJpegPath.empty())
                {
                    std::vector<uint8_t> sidecarBytes;
                    std::wstring sidecarStatus;
                    const HRESULT sidecarReadHr = ReadFileAllBytes(fileSystem.get(), item.sidecarJpegPath, sidecarBytes, sidecarStatus);
                    if (SUCCEEDED(sidecarReadHr) && ! sidecarBytes.empty())
                    {
                        const HRESULT sidecarDecodeHr = DecodeJpegToBgraTurboJpeg(sidecarBytes.data(), sidecarBytes.size(), thumbW, thumbH, thumbBgra);
                        if (SUCCEEDED(sidecarDecodeHr))
                        {
                            thumbDecoded = true;
                            thumbSource  = CachedImage::ThumbSource::SidecarJpeg;
                            thumbExif    = ExtractExifFromJpeg(sidecarBytes.data(), sidecarBytes.size());
                        }
                    }
                }

                if (! thumbDecoded)
                {
                    thumbDecoded = TryDecodeRawEmbeddedThumbnailFromBufferToBgra(fileBytes, thumbW, thumbH, thumbBgra, thumbAvailable, thumbExif);
                    if (thumbDecoded)
                    {
                        thumbSource = CachedImage::ThumbSource::Embedded;
                    }
                }

                const bool needRawDecode = (prefetchMode == DisplayMode::Raw) || ! thumbDecoded;
                if (needRawDecode && ! hasRawAlready)
                {
                    rawHr = DecodeRawFullImageFromBufferToBgra(decodeCfg, fileBytes, rawW, rawH, rawBgra, rawStatus, rawExif, nullptr);
                }
            }

            if (requestId == _openRequestId.load(std::memory_order_acquire))
            {
                std::unique_ptr<CachedImage> created(new (std::nothrow) CachedImage{});

                std::scoped_lock lock(_cacheMutex);
                auto& entry = _imageCache[path];
                if (! entry)
                {
                    entry = std::move(created);
                }

                if (entry)
                {
                    entry->thumbAvailable = entry->thumbAvailable || thumbAvailable;

                    if (thumbDecoded && thumbW != 0 && thumbH != 0 && ! thumbBgra.empty())
                    {
                        entry->thumbWidth       = thumbW;
                        entry->thumbHeight      = thumbH;
                        entry->thumbOrientation = NormalizeExifOrientation(thumbExif.orientation);
                        entry->thumbBgra        = std::move(thumbBgra);
                        entry->thumbDecoded     = true;
                        entry->thumbSource      = thumbSource;
                    }

                    if (SUCCEEDED(rawHr) && rawW != 0 && rawH != 0 && ! rawBgra.empty())
                    {
                        entry->rawWidth       = rawW;
                        entry->rawHeight      = rawH;
                        entry->rawOrientation = NormalizeExifOrientation(rawExif.orientation);
                        entry->rawBgra        = std::move(rawBgra);
                    }

                    if (rawExif.valid)
                    {
                        entry->exif.camera         = rawExif.camera;
                        entry->exif.lens           = rawExif.lens;
                        entry->exif.dateTime       = rawExif.dateTime;
                        entry->exif.iso            = rawExif.iso;
                        entry->exif.shutterSeconds = rawExif.shutterSeconds;
                        entry->exif.aperture       = rawExif.aperture;
                        entry->exif.focalLengthMm  = rawExif.focalLengthMm;
                        entry->exif.valid          = true;
                    }
                    else if (! entry->exif.valid && thumbExif.valid)
                    {
                        entry->exif.camera         = thumbExif.camera;
                        entry->exif.lens           = thumbExif.lens;
                        entry->exif.dateTime       = thumbExif.dateTime;
                        entry->exif.iso            = thumbExif.iso;
                        entry->exif.shutterSeconds = thumbExif.shutterSeconds;
                        entry->exif.aperture       = thumbExif.aperture;
                        entry->exif.focalLengthMm  = thumbExif.focalLengthMm;
                        entry->exif.valid          = true;
                    }
                }
            }

            std::scoped_lock lock(_cacheMutex);
            _inflightDecodes.erase(path);
        }
    };

    const BOOL queued = TrySubmitThreadpoolCallback(
        [](PTP_CALLBACK_INSTANCE /*instance*/, void* context) noexcept
        {
            std::unique_ptr<PrefetchWorkItem> ctx(static_cast<PrefetchWorkItem*>(context));
            if (! ctx)
            {
                return;
            }

            static_cast<void>(ctx->moduleKeepAlive);
            if (ctx->work)
            {
                ctx->work();
            }
        },
        ctx.get(),
        nullptr);

    if (queued == 0)
    {
        Debug::Error(L"ViewerImgRaw: Failed to queue neighbor prefetch work item.");
    }

    ctx.release();
}

void ViewerImgRaw::StartAsyncOpen(HWND hwnd, std::wstring_view path, bool updateOtherFiles) noexcept
{
    if (! hwnd || path.empty())
    {
        return;
    }

    const bool samePath      = (_currentPath == path);
    const uint64_t requestId = _openRequestId.fetch_add(1, std::memory_order_acq_rel) + 1;

    EndLoadingUi();

    _statusMessage.clear();
    _imageBitmap.reset();
    _currentImageOwned.reset();
    _currentImage = nullptr;
    _currentImageKey.clear();
    _exifOverlayText.clear();
    _displayedMode      = _displayMode;
    _rawProgressPercent = -1;
    _rawProgressStage   = -1;
    _rawProgressStageText.clear();
    _panOffsetXPx = 0.0f;
    _panOffsetYPx = 0.0f;
    _panning      = false;
    UpdateScrollBars(hwnd);
    _baseOrientation = 1;
    _viewOrientation = 1;
    if (! samePath)
    {
        _userOrientation         = 1;
        _orientationUserModified = false;
    }
    else
    {
        _orientationUserModified = (_userOrientation != 1);
    }

    _currentPath.assign(path);
    if (updateOtherFiles)
    {
        _otherItems.clear();
        OtherItem item{};
        item.primaryPath = _currentPath;
        item.label       = LeafNameFromPath(_currentPath);
        if (item.label.empty())
        {
            item.label = _currentPath;
        }
        const std::wstring extLower = ToLowerCopy(PathExtensionView(_currentPath));
        item.isRaw                  = IsLikelyRawExtension(extLower) && ! IsWicImageExtension(extLower);

        _otherItems.push_back(item);
        _otherIndex = 0;
        _currentSidecarJpegPath.clear();
        _currentLabel = item.label;
        RefreshFileCombo(hwnd);
    }
    else
    {
        SyncFileComboSelection();
    }

    if (_currentLabel.empty())
    {
        _currentLabel = LeafNameFromPath(_currentPath);
        if (_currentLabel.empty())
        {
            _currentLabel = _currentPath;
        }
    }

    const std::wstring title = FormatStringResource(g_hInstance, IDS_VIEWERRAW_TITLE_FORMAT, _currentLabel);
    if (! title.empty())
    {
        SetWindowTextW(hwnd, title.c_str());
    }

    bool continueDecoding = false;
    if (TryUseCachedImage(hwnd, _currentPath, continueDecoding))
    {
        UpdateScrollBars(hwnd);
        InvalidateRect(hwnd, &_contentRect, FALSE);
        InvalidateRect(hwnd, &_statusRect, FALSE);

        if (! continueDecoding)
        {
            UpdateNeighborCache(requestId);
            return;
        }
    }

    _isLoading = true;
    BeginLoadingUi();
    InvalidateRect(hwnd, &_contentRect, FALSE);
    InvalidateRect(hwnd, &_statusRect, FALSE);

    const wil::com_ptr<IFileSystem> fileSystem = _fileSystem;
    const std::wstring pathCopy(_currentPath);
    const std::wstring labelCopy(_currentLabel);
    const std::wstring sidecarPathCopy(_currentSidecarJpegPath);

    const bool decodeRawOnly      = ToLowerCopy(PathExtensionView(pathCopy)) == L".tif";
    const DisplayMode desiredMode = _displayMode;

    RawDecodeSettings decodeCfg{};
    decodeCfg.halfSize    = _config.halfSize;
    decodeCfg.useCameraWb = _config.useCameraWb;
    decodeCfg.autoWb      = _config.autoWb;

    const uint32_t cfgSignature = static_cast<uint32_t>((_config.halfSize ? 1u : 0u) | (_config.useCameraWb ? 2u : 0u) | (_config.autoWb ? 4u : 0u));

    AddRef();

    struct AsyncOpenWorkItem final
    {
        AsyncOpenWorkItem()                                    = default;
        AsyncOpenWorkItem(const AsyncOpenWorkItem&)            = delete;
        AsyncOpenWorkItem& operator=(const AsyncOpenWorkItem&) = delete;

        wil::unique_hmodule moduleKeepAlive;
        std::function<void()> work;
    };

    auto ctx = std::unique_ptr<AsyncOpenWorkItem>(new (std::nothrow) AsyncOpenWorkItem{});

    ctx->moduleKeepAlive = AcquireModuleReferenceFromAddress(&kViewerImgRawModuleAnchor);
    ctx->work            = [this,
                 fileSystem,
                 hwnd,
                 requestId,
                 path = pathCopy,
                 updateOtherFiles,
                 cfgSignature,
                 label       = labelCopy,
                 sidecarPath = sidecarPathCopy,
                 desiredMode,
                 decodeCfg,
                 decodeRawOnly]() mutable
    {
        auto releaseSelf = wil::scope_exit([&] { Release(); });

        std::unique_ptr<AsyncOpenResult> result(new (std::nothrow) AsyncOpenResult{});
        if (! result)
        {
            return;
        }

        auto postResult = wil::scope_exit(
            [&]
            {
                if (! hwnd || GetWindowLongPtrW(hwnd, GWLP_USERDATA) != reinterpret_cast<LONG_PTR>(this))
                {
                    return;
                }

                static_cast<void>(PostMessagePayload(hwnd, kAsyncOpenCompleteMessage, 0, std::move(result)));
            });

        result->viewer           = this;
        result->requestId        = requestId;
        result->path             = path;
        result->updateOtherFiles = updateOtherFiles;
        result->configSignature  = cfgSignature;
        result->frameMode        = desiredMode;
        result->isFinal          = true;

        if (! fileSystem)
        {
            result->hr            = E_FAIL;
            result->statusMessage = L"ViewerImgRaw: Active filesystem is missing.";
            return;
        }

        {
            std::scoped_lock lock(_cacheMutex);
            _inflightDecodes.insert(path);
        }

        if (desiredMode == DisplayMode::Thumbnail && ! sidecarPath.empty() && IsLikelyRawExtension(ToLowerCopy(PathExtensionView(path))))
        {
            std::vector<uint8_t> sidecarBytes;
            std::wstring sidecarStatus;
            const HRESULT sidecarReadHr = ReadFileAllBytes(fileSystem.get(), sidecarPath, sidecarBytes, sidecarStatus);
            if (SUCCEEDED(sidecarReadHr) && ! sidecarBytes.empty())
            {
                static constexpr int kJpegPreviewMaxDim = 2048;
                const ExifData sidecarExif              = ExtractExifFromJpeg(sidecarBytes.data(), sidecarBytes.size());
                const bool doPreview                    = ShouldRenderJpegProgressively(sidecarBytes.data(), sidecarBytes.size(), kJpegPreviewMaxDim);

                if (doPreview && requestId == _openRequestId.load(std::memory_order_acquire))
                {
                    uint32_t previewW = 0;
                    uint32_t previewH = 0;
                    std::vector<uint8_t> previewBgra;
                    const HRESULT previewHr =
                        DecodeJpegToBgraTurboJpegScaled(sidecarBytes.data(), sidecarBytes.size(), kJpegPreviewMaxDim, previewW, previewH, previewBgra);
                    if (SUCCEEDED(previewHr) && previewW != 0 && previewH != 0 && ! previewBgra.empty())
                    {
                        std::unique_ptr<AsyncOpenResult> preview(new (std::nothrow) AsyncOpenResult{});
                        if (preview)
                        {
                            preview->viewer           = this;
                            preview->requestId        = requestId;
                            preview->hr               = S_OK;
                            preview->path             = path;
                            preview->updateOtherFiles = updateOtherFiles;
                            preview->configSignature  = cfgSignature;
                            preview->frameMode        = DisplayMode::Thumbnail;
                            preview->isFinal          = false;
                            preview->thumbAvailable   = true;
                            preview->thumbSource      = CachedImage::ThumbSource::SidecarJpeg;
                            preview->width            = previewW;
                            preview->height           = previewH;
                            preview->bgra             = std::move(previewBgra);
                            preview->statusMessage.clear();

                            preview->exif.orientation = sidecarExif.orientation;
                            if (sidecarExif.valid)
                            {
                                preview->exif.camera         = sidecarExif.camera;
                                preview->exif.lens           = sidecarExif.lens;
                                preview->exif.dateTime       = sidecarExif.dateTime;
                                preview->exif.iso            = sidecarExif.iso;
                                preview->exif.shutterSeconds = sidecarExif.shutterSeconds;
                                preview->exif.aperture       = sidecarExif.aperture;
                                preview->exif.focalLengthMm  = sidecarExif.focalLengthMm;
                                preview->exif.valid          = true;
                            }
                            else
                            {
                                preview->exif.valid = false;
                            }

                            if (hwnd && GetWindowLongPtrW(hwnd, GWLP_USERDATA) == reinterpret_cast<LONG_PTR>(this))
                            {
                                static_cast<void>(PostMessagePayload(hwnd, kAsyncOpenCompleteMessage, 0, std::move(preview)));
                            }
                        }
                    }
                }

                uint32_t thumbW = 0;
                uint32_t thumbH = 0;
                std::vector<uint8_t> thumbBgra;
                const HRESULT sidecarDecodeHr = DecodeJpegToBgraTurboJpeg(sidecarBytes.data(), sidecarBytes.size(), thumbW, thumbH, thumbBgra);
                if (SUCCEEDED(sidecarDecodeHr) && thumbW != 0 && thumbH != 0 && ! thumbBgra.empty())
                {
                    result->exif.orientation = sidecarExif.orientation;
                    if (sidecarExif.valid)
                    {
                        result->exif.camera         = sidecarExif.camera;
                        result->exif.lens           = sidecarExif.lens;
                        result->exif.dateTime       = sidecarExif.dateTime;
                        result->exif.iso            = sidecarExif.iso;
                        result->exif.shutterSeconds = sidecarExif.shutterSeconds;
                        result->exif.aperture       = sidecarExif.aperture;
                        result->exif.focalLengthMm  = sidecarExif.focalLengthMm;
                        result->exif.valid          = true;
                    }
                    else
                    {
                        result->exif.valid = false;
                    }

                    result->hr             = S_OK;
                    result->frameMode      = DisplayMode::Thumbnail;
                    result->isFinal        = true;
                    result->thumbAvailable = true;
                    result->thumbSource    = CachedImage::ThumbSource::SidecarJpeg;
                    result->width          = thumbW;
                    result->height         = thumbH;
                    result->bgra           = std::move(thumbBgra);
                    result->statusMessage.clear();
                    return;
                }
            }
        }

        std::vector<uint8_t> fileBytes;
        std::wstring readStatus;
        const HRESULT readHr = ReadFileAllBytes(fileSystem.get(), path, fileBytes, readStatus);
        if (FAILED(readHr))
        {
            result->hr            = readHr;
            result->statusMessage = std::move(readStatus);
        }
        else
        {
            const std::wstring extLower = ToLowerCopy(PathExtensionView(path));

            if (IsWicImageExtension(extLower))
            {
                const bool isJpeg = IsJpegExtension(extLower);
                if (isJpeg)
                {
                    static constexpr int kJpegPreviewMaxDim = 2048;
                    const ExifData jpegExif                 = ExtractExifFromJpeg(fileBytes.data(), fileBytes.size());

                    const bool doPreview = ShouldRenderJpegProgressively(fileBytes.data(), fileBytes.size(), kJpegPreviewMaxDim);
                    if (doPreview && requestId == _openRequestId.load(std::memory_order_acquire))
                    {
                        uint32_t previewW = 0;
                        uint32_t previewH = 0;
                        std::vector<uint8_t> previewBgra;
                        const HRESULT previewHr =
                            DecodeJpegToBgraTurboJpegScaled(fileBytes.data(), fileBytes.size(), kJpegPreviewMaxDim, previewW, previewH, previewBgra);
                        if (SUCCEEDED(previewHr) && previewW != 0 && previewH != 0 && ! previewBgra.empty())
                        {
                            std::unique_ptr<AsyncOpenResult> preview(new (std::nothrow) AsyncOpenResult{});
                            if (preview)
                            {
                                preview->viewer           = this;
                                preview->requestId        = requestId;
                                preview->hr               = S_OK;
                                preview->path             = path;
                                preview->updateOtherFiles = updateOtherFiles;
                                preview->configSignature  = cfgSignature;
                                preview->frameMode        = DisplayMode::Raw;
                                preview->isFinal          = false;
                                preview->thumbAvailable   = false;
                                preview->thumbSource      = CachedImage::ThumbSource::None;
                                preview->width            = previewW;
                                preview->height           = previewH;
                                preview->bgra             = std::move(previewBgra);
                                preview->statusMessage.clear();

                                preview->exif.orientation = jpegExif.orientation;
                                if (jpegExif.valid)
                                {
                                    preview->exif.camera         = jpegExif.camera;
                                    preview->exif.lens           = jpegExif.lens;
                                    preview->exif.dateTime       = jpegExif.dateTime;
                                    preview->exif.iso            = jpegExif.iso;
                                    preview->exif.shutterSeconds = jpegExif.shutterSeconds;
                                    preview->exif.aperture       = jpegExif.aperture;
                                    preview->exif.focalLengthMm  = jpegExif.focalLengthMm;
                                    preview->exif.valid          = true;
                                }
                                else
                                {
                                    preview->exif.valid = false;
                                }

                                if (hwnd && GetWindowLongPtrW(hwnd, GWLP_USERDATA) == reinterpret_cast<LONG_PTR>(this))
                                {
                                    static_cast<void>(PostMessagePayload(hwnd, kAsyncOpenCompleteMessage, 0, std::move(preview)));
                                }
                            }
                        }
                    }

                    uint32_t w = 0;
                    uint32_t h = 0;
                    std::vector<uint8_t> bgra;
                    const HRESULT decodeHr = DecodeJpegToBgraTurboJpeg(fileBytes.data(), fileBytes.size(), w, h, bgra);
                    if (SUCCEEDED(decodeHr))
                    {
                        result->exif.orientation = jpegExif.orientation;
                        if (jpegExif.valid)
                        {
                            result->exif.camera         = jpegExif.camera;
                            result->exif.lens           = jpegExif.lens;
                            result->exif.dateTime       = jpegExif.dateTime;
                            result->exif.iso            = jpegExif.iso;
                            result->exif.shutterSeconds = jpegExif.shutterSeconds;
                            result->exif.aperture       = jpegExif.aperture;
                            result->exif.focalLengthMm  = jpegExif.focalLengthMm;
                            result->exif.valid          = true;
                        }
                        else
                        {
                            result->exif.valid = false;
                        }

                        result->hr        = S_OK;
                        result->frameMode = DisplayMode::Raw;
                        result->isFinal   = true;
                        result->width     = w;
                        result->height    = h;
                        result->bgra      = std::move(bgra);
                        result->statusMessage.clear();
                        result->thumbAvailable = false;
                        result->thumbSource    = CachedImage::ThumbSource::None;
                        return;
                    }
                }
                else
                {
                    uint32_t w = 0;
                    uint32_t h = 0;
                    std::vector<uint8_t> bgra;
                    const HRESULT decodeHr = DecodeImageToBgraWic(fileBytes.data(), fileBytes.size(), w, h, bgra);
                    if (SUCCEEDED(decodeHr))
                    {
                        result->exif.orientation = 1;
                        result->exif.valid       = false;

                        result->hr        = S_OK;
                        result->frameMode = DisplayMode::Raw;
                        result->isFinal   = true;
                        result->width     = w;
                        result->height    = h;
                        result->bgra      = std::move(bgra);
                        result->statusMessage.clear();
                        result->thumbAvailable = false;
                        result->thumbSource    = CachedImage::ThumbSource::None;
                        return;
                    }
                }
            }

            const bool isRawInput = IsLikelyRawExtension(extLower) && ! IsWicImageExtension(extLower);
            if (! isRawInput)
            {
                result->hr            = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                result->statusMessage = L"ViewerImgRaw: Unsupported file format.";
                return;
            }

            bool thumbAvailable = decodeRawOnly || ! sidecarPath.empty();
            ExifData thumbExif{};
            bool thumbDecoded                    = false;
            CachedImage::ThumbSource thumbSource = CachedImage::ThumbSource::None;
            uint32_t thumbW                      = 0;
            uint32_t thumbH                      = 0;
            std::vector<uint8_t> thumbBgra;

            if (! decodeRawOnly)
            {
                if (! sidecarPath.empty())
                {
                    std::vector<uint8_t> sidecarBytes;
                    std::wstring sidecarStatus;
                    const HRESULT sidecarReadHr = ReadFileAllBytes(fileSystem.get(), sidecarPath, sidecarBytes, sidecarStatus);
                    if (SUCCEEDED(sidecarReadHr) && ! sidecarBytes.empty())
                    {
                        const HRESULT sidecarDecodeHr = DecodeJpegToBgraTurboJpeg(sidecarBytes.data(), sidecarBytes.size(), thumbW, thumbH, thumbBgra);
                        if (SUCCEEDED(sidecarDecodeHr))
                        {
                            thumbDecoded = true;
                            thumbSource  = CachedImage::ThumbSource::SidecarJpeg;
                            thumbExif    = ExtractExifFromJpeg(sidecarBytes.data(), sidecarBytes.size());
                        }
                    }
                }

                if (! thumbDecoded)
                {
                    thumbDecoded = TryDecodeRawEmbeddedThumbnailFromBufferToBgra(fileBytes, thumbW, thumbH, thumbBgra, thumbAvailable, thumbExif);
                    if (thumbDecoded)
                    {
                        thumbSource = CachedImage::ThumbSource::Embedded;
                    }
                }

                if (thumbDecoded && requestId == _openRequestId.load(std::memory_order_acquire))
                {
                    std::unique_ptr<AsyncOpenResult> preview(new (std::nothrow) AsyncOpenResult{});
                    if (preview)
                    {
                        preview->viewer           = this;
                        preview->requestId        = requestId;
                        preview->hr               = S_OK;
                        preview->path             = path;
                        preview->updateOtherFiles = updateOtherFiles;
                        preview->configSignature  = cfgSignature;
                        preview->frameMode        = DisplayMode::Thumbnail;
                        preview->isFinal          = false;
                        preview->thumbAvailable   = true;
                        preview->thumbSource      = thumbSource;
                        preview->width            = thumbW;
                        preview->height           = thumbH;
                        preview->bgra             = std::move(thumbBgra);
                        preview->exif.orientation = thumbExif.orientation;

                        if (thumbExif.valid)
                        {
                            preview->exif.camera         = thumbExif.camera;
                            preview->exif.lens           = thumbExif.lens;
                            preview->exif.dateTime       = thumbExif.dateTime;
                            preview->exif.iso            = thumbExif.iso;
                            preview->exif.shutterSeconds = thumbExif.shutterSeconds;
                            preview->exif.aperture       = thumbExif.aperture;
                            preview->exif.focalLengthMm  = thumbExif.focalLengthMm;
                            preview->exif.valid          = true;
                        }

                        if (hwnd && GetWindowLongPtrW(hwnd, GWLP_USERDATA) == reinterpret_cast<LONG_PTR>(this))
                        {
                            static_cast<void>(PostMessagePayload(hwnd, kAsyncOpenCompleteMessage, 0, std::move(preview)));
                        }

                        if (desiredMode == DisplayMode::Thumbnail)
                        {
                            return;
                        }
                    }
                }
            }

            const bool needRawDecode = (desiredMode == DisplayMode::Raw) || ! thumbDecoded;
            if (needRawDecode && requestId == _openRequestId.load(std::memory_order_acquire))
            {
                ExifData rawExif{};

                LibRawProgressHost host{};
                host.requestIdCounter = &_openRequestId;
                host.requestId        = requestId;
                host.hwnd             = hwnd;

                result->thumbAvailable = thumbAvailable;
                result->frameMode      = DisplayMode::Raw;
                result->isFinal        = true;

                result->hr = DecodeRawFullImageFromBufferToBgra(
                    decodeCfg, fileBytes, result->width, result->height, result->bgra, result->statusMessage, rawExif, &host);

                if (FAILED(result->hr))
                {
                    uint32_t w = 0;
                    uint32_t h = 0;
                    std::vector<uint8_t> bgra;
                    const HRESULT fallbackHr = IsJpegExtension(extLower) ? DecodeJpegToBgraTurboJpeg(fileBytes.data(), fileBytes.size(), w, h, bgra)
                                                                         : DecodeImageToBgraWic(fileBytes.data(), fileBytes.size(), w, h, bgra);
                    if (SUCCEEDED(fallbackHr))
                    {
                        result->hr        = S_OK;
                        result->frameMode = DisplayMode::Raw;
                        result->width     = w;
                        result->height    = h;
                        result->bgra      = std::move(bgra);
                        result->statusMessage.clear();
                        rawExif = {};
                    }
                }

                const ExifData& effectiveExif = rawExif.valid ? rawExif : thumbExif;
                result->exif.camera           = effectiveExif.camera;
                result->exif.lens             = effectiveExif.lens;
                result->exif.dateTime         = effectiveExif.dateTime;
                result->exif.iso              = effectiveExif.iso;
                result->exif.shutterSeconds   = effectiveExif.shutterSeconds;
                result->exif.aperture         = effectiveExif.aperture;
                result->exif.focalLengthMm    = effectiveExif.focalLengthMm;
                result->exif.orientation      = 1;
                result->exif.valid            = effectiveExif.valid;
            }
        }
    };

    const BOOL queued = TrySubmitThreadpoolCallback(
        [](PTP_CALLBACK_INSTANCE /*instance*/, void* context) noexcept
        {
            std::unique_ptr<AsyncOpenWorkItem> ctx(static_cast<AsyncOpenWorkItem*>(context));
            if (! ctx)
            {
                return;
            }

            static_cast<void>(ctx->moduleKeepAlive);
            if (ctx->work)
            {
                ctx->work();
            }
        },
        ctx.get(),
        nullptr);

    if (queued == 0)
    {
        Debug::Error(L"ViewerImgRaw: Failed to queue async open work item.");
    }

    ctx.release();
}

void ViewerImgRaw::OnAsyncOpenComplete(std::unique_ptr<AsyncOpenResult> result) noexcept
{
    if (! result)
    {
        return;
    }

    if (result->requestId != _openRequestId.load(std::memory_order_acquire))
    {
        return;
    }

    if (result->configSignature != static_cast<uint32_t>((_config.halfSize ? 1u : 0u) | (_config.useCameraWb ? 2u : 0u) | (_config.autoWb ? 4u : 0u)))
    {
        return;
    }

    {
        std::scoped_lock lock(_cacheMutex);
        _inflightDecodes.erase(result->path);
    }

    if (_isLoading && result->isFinal)
    {
        EndLoadingUi();
    }

    if (SUCCEEDED(result->hr))
    {
        const bool cachingEnabled = (_config.prevCache > 0 || _config.nextCache > 0);
        CachedImage* image        = nullptr;

        if (cachingEnabled)
        {
            std::unique_ptr<CachedImage> created(new (std::nothrow) CachedImage{});
            std::scoped_lock lock(_cacheMutex);
            auto& entry = _imageCache[result->path];
            if (! entry)
            {
                entry = std::move(created);
            }
            image = entry.get();
        }
        else
        {
            if (! _currentImageOwned)
            {
                _currentImageOwned.reset(new (std::nothrow) CachedImage{});
            }
            image = _currentImageOwned.get();
        }

        if (! image)
        {
            result->hr            = E_OUTOFMEMORY;
            result->statusMessage = L"ViewerImgRaw: Out of memory while storing decoded image.";
        }
        else
        {
            if (result->thumbAvailable)
            {
                image->thumbAvailable = true;
            }

            if (result->frameMode == DisplayMode::Thumbnail)
            {
                image->thumbOrientation = NormalizeExifOrientation(result->exif.orientation);
                if (result->width != 0 && result->height != 0 && ! result->bgra.empty())
                {
                    image->thumbWidth   = result->width;
                    image->thumbHeight  = result->height;
                    image->thumbBgra    = std::move(result->bgra);
                    image->thumbDecoded = true;
                    image->thumbSource  = result->thumbSource;
                }
                _displayedMode = DisplayMode::Thumbnail;
            }
            else
            {
                image->rawOrientation = NormalizeExifOrientation(result->exif.orientation);
                if (result->width != 0 && result->height != 0 && ! result->bgra.empty())
                {
                    image->rawWidth  = result->width;
                    image->rawHeight = result->height;
                    image->rawBgra   = std::move(result->bgra);
                }
                _displayedMode = DisplayMode::Raw;
            }

            if (result->exif.valid)
            {
                image->exif = result->exif;
            }

            if (cachingEnabled)
            {
                _currentImageOwned.reset();
            }
            _currentImage    = image;
            _currentImageKey = result->path;

            _statusMessage.clear();
            _imageBitmap.reset();
            UpdateOrientationState();
            RebuildExifOverlayText();

            if (_hostAlerts)
            {
                static_cast<void>(_hostAlerts->ClearAlert(HOST_ALERT_SCOPE_WINDOW, nullptr));
            }
            _alertVisible = false;

            if (result->isFinal)
            {
                UpdateNeighborCache(result->requestId);
            }
        }
    }

    if (FAILED(result->hr))
    {
        const bool keepImage = HasDisplayImage();
        if (! keepImage)
        {
            _currentImageOwned.reset();
            _currentImage = nullptr;
            _currentImageKey.clear();
        }
        _imageBitmap.reset();
        _statusMessage =
            keepImage ? L"" : (result->statusMessage.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERRAW_STATUS_ERROR) : result->statusMessage);

        if (_hostAlerts)
        {
            const std::wstring message = result->statusMessage.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERRAW_STATUS_ERROR) : result->statusMessage;
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
        }
        _alertVisible = true;
    }

    if (_hWnd)
    {
        UpdateMenuChecks(_hWnd.get());
        UpdateScrollBars(_hWnd.get());
        InvalidateRect(_hWnd.get(), &_contentRect, FALSE);
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
}

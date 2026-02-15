#pragma once

#include <algorithm>
#include <cwctype>
#include <string>
#include <string_view>

[[nodiscard]] inline bool EqualsIgnoreCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    for (size_t i = 0; i < a.size(); ++i)
    {
        if (std::towlower(a[i]) != std::towlower(b[i]))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] inline std::wstring_view PathExtensionView(std::wstring_view path) noexcept
{
    const size_t dot = path.find_last_of(L'.');
    if (dot == std::wstring_view::npos || dot + 1 >= path.size())
    {
        return {};
    }

    const size_t slash = path.find_last_of(L"/\\");
    if (slash != std::wstring_view::npos && dot < slash)
    {
        return {};
    }

    return path.substr(dot);
}

[[nodiscard]] inline std::wstring_view PathWithoutExtensionView(std::wstring_view path) noexcept
{
    const size_t dot = path.find_last_of(L'.');
    if (dot == std::wstring_view::npos || dot == 0)
    {
        return path;
    }

    const size_t slash = path.find_last_of(L"/\\");
    if (slash != std::wstring_view::npos && dot < slash)
    {
        return path;
    }

    return path.substr(0, dot);
}

[[nodiscard]] inline std::wstring ToLowerCopy(std::wstring_view text) noexcept
{
    std::wstring out;
    out.assign(text);
    for (wchar_t& ch : out)
    {
        ch = static_cast<wchar_t>(std::towlower(ch));
    }
    return out;
}

[[nodiscard]] inline std::wstring LeafNameFromPath(std::wstring_view path) noexcept
{
    const size_t slash = path.find_last_of(L"/\\");
    if (slash == std::wstring_view::npos)
    {
        return std::wstring(path);
    }
    return std::wstring(path.substr(slash + 1));
}

[[nodiscard]] inline bool IsJpegExtension(std::wstring_view extLower) noexcept
{
    return EqualsIgnoreCase(extLower, L".jpg") || EqualsIgnoreCase(extLower, L".jpeg") || EqualsIgnoreCase(extLower, L".jpe");
}

[[nodiscard]] inline bool IsWicImageExtension(std::wstring_view extLower) noexcept
{
    return EqualsIgnoreCase(extLower, L".bmp") || EqualsIgnoreCase(extLower, L".dib") || EqualsIgnoreCase(extLower, L".gif") ||
           EqualsIgnoreCase(extLower, L".ico") || EqualsIgnoreCase(extLower, L".jpe") || EqualsIgnoreCase(extLower, L".jpeg") ||
           EqualsIgnoreCase(extLower, L".jpg") || EqualsIgnoreCase(extLower, L".png") || EqualsIgnoreCase(extLower, L".tif") ||
           EqualsIgnoreCase(extLower, L".tiff") || EqualsIgnoreCase(extLower, L".wdp") || EqualsIgnoreCase(extLower, L".jxr") ||
           EqualsIgnoreCase(extLower, L".hdp");
}

[[nodiscard]] inline bool IsLikelyRawExtension(std::wstring_view extLower) noexcept
{
    return EqualsIgnoreCase(extLower, L".3fr") || EqualsIgnoreCase(extLower, L".ari") || EqualsIgnoreCase(extLower, L".arw") ||
           EqualsIgnoreCase(extLower, L".bay") || EqualsIgnoreCase(extLower, L".braw") || EqualsIgnoreCase(extLower, L".crw") ||
           EqualsIgnoreCase(extLower, L".cr2") || EqualsIgnoreCase(extLower, L".cr3") || EqualsIgnoreCase(extLower, L".cap") ||
           EqualsIgnoreCase(extLower, L".data") || EqualsIgnoreCase(extLower, L".dcs") || EqualsIgnoreCase(extLower, L".dcr") ||
           EqualsIgnoreCase(extLower, L".dng") || EqualsIgnoreCase(extLower, L".drf") || EqualsIgnoreCase(extLower, L".eip") ||
           EqualsIgnoreCase(extLower, L".erf") || EqualsIgnoreCase(extLower, L".fff") || EqualsIgnoreCase(extLower, L".gpr") ||
           EqualsIgnoreCase(extLower, L".iiq") || EqualsIgnoreCase(extLower, L".k25") || EqualsIgnoreCase(extLower, L".kdc") ||
           EqualsIgnoreCase(extLower, L".mdc") || EqualsIgnoreCase(extLower, L".mef") || EqualsIgnoreCase(extLower, L".mos") ||
           EqualsIgnoreCase(extLower, L".mrw") || EqualsIgnoreCase(extLower, L".nef") || EqualsIgnoreCase(extLower, L".nrw") ||
           EqualsIgnoreCase(extLower, L".obm") || EqualsIgnoreCase(extLower, L".orf") || EqualsIgnoreCase(extLower, L".pef") ||
           EqualsIgnoreCase(extLower, L".ptx") || EqualsIgnoreCase(extLower, L".pxn") || EqualsIgnoreCase(extLower, L".r3d") ||
           EqualsIgnoreCase(extLower, L".raf") || EqualsIgnoreCase(extLower, L".raw") || EqualsIgnoreCase(extLower, L".rwl") ||
           EqualsIgnoreCase(extLower, L".rw2") || EqualsIgnoreCase(extLower, L".rwz") || EqualsIgnoreCase(extLower, L".sr2") ||
           EqualsIgnoreCase(extLower, L".srf") || EqualsIgnoreCase(extLower, L".srw") || EqualsIgnoreCase(extLower, L".tif") ||
           EqualsIgnoreCase(extLower, L".x3f");
}

struct ExifOrientationLin final
{
    int m11 = 1;
    int m12 = 0;
    int m21 = 0;
    int m22 = 1;
};

[[nodiscard]] constexpr ExifOrientationLin LinFromExifOrientation(uint16_t orientation) noexcept
{
    switch (orientation)
    {
        case 2: return ExifOrientationLin{-1, 0, 0, 1};
        case 3: return ExifOrientationLin{-1, 0, 0, -1};
        case 4: return ExifOrientationLin{1, 0, 0, -1};
        case 5: return ExifOrientationLin{0, 1, 1, 0};
        case 6: return ExifOrientationLin{0, 1, -1, 0};
        case 7: return ExifOrientationLin{0, -1, -1, 0};
        case 8: return ExifOrientationLin{0, -1, 1, 0};
        case 1:
        default: return ExifOrientationLin{1, 0, 0, 1};
    }
}

[[nodiscard]] constexpr ExifOrientationLin MultiplyExifOrientation(const ExifOrientationLin& a, const ExifOrientationLin& b) noexcept
{
    ExifOrientationLin out{};
    out.m11 = a.m11 * b.m11 + a.m21 * b.m12;
    out.m12 = a.m12 * b.m11 + a.m22 * b.m12;
    out.m21 = a.m11 * b.m21 + a.m21 * b.m22;
    out.m22 = a.m12 * b.m21 + a.m22 * b.m22;
    return out;
}

[[nodiscard]] constexpr uint16_t ExifOrientationFromLin(const ExifOrientationLin& lin) noexcept
{
    for (uint16_t o = 1; o <= 8; ++o)
    {
        const ExifOrientationLin cand = LinFromExifOrientation(o);
        if (cand.m11 == lin.m11 && cand.m12 == lin.m12 && cand.m21 == lin.m21 && cand.m22 == lin.m22)
        {
            return o;
        }
    }
    return 1;
}

[[nodiscard]] inline uint16_t NormalizeExifOrientation(uint16_t orientation) noexcept
{
    return (orientation >= 1 && orientation <= 8) ? orientation : static_cast<uint16_t>(1);
}

// Returns EXIF orientation corresponding to applying `a` after `b` (composition: a ∘ b).
[[nodiscard]] inline uint16_t ComposeExifOrientation(uint16_t a, uint16_t b) noexcept
{
    const ExifOrientationLin la = LinFromExifOrientation(NormalizeExifOrientation(a));
    const ExifOrientationLin lb = LinFromExifOrientation(NormalizeExifOrientation(b));
    return ExifOrientationFromLin(MultiplyExifOrientation(la, lb));
}

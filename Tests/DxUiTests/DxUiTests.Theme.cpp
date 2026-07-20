#include "DxUiTestHelpers.h"
#include "DxUiThemePalette.h"
#include "FileMetadataFormatting.h"
#include "HandleIo.h"
#include "PaginationGuard.h"
#include "PathUtils.h"
#include "SettingsStore.h"
#include "ThroughputParsing.h"
#include "UriEncoding.h"
#include "ViewerFileComboHost.h"
#include "ViewerTitleBarTheme.h"
#include "WindowSizing.h"
#include "YyjsonHelpers.h"

namespace
{

void TestColorFromArgbPreservesAlpha()
{
    using namespace RedSalamander::DxUi;

    RequireColorNear(ColorFromArgb(0x80402010u),
                     D2D1::ColorF(0x40 / 255.0f, 0x20 / 255.0f, 0x10 / 255.0f, 0x80 / 255.0f),
                     "shared ARGB conversion preserves alpha and channel order");
}

void TestSharedColorRefArgbAndTruncatingBlendGoldenValues()
{
    using namespace Common::Colors;

    Require(ColorRefFromArgb(0x7F123456u) == RGB(0x12, 0x34, 0x56), "ARGB-to-COLORREF conversion drops alpha and preserves RGB byte order");
    Require(BlendColorRefTruncate(RGB(10, 20, 30), RGB(110, 220, 230), 128u) == RGB(60, 120, 130),
            "COLORREF blend preserves the existing truncating 8-bit channel policy");
    Require(BlendColorRefTruncate(RGB(1, 2, 3), RGB(200, 201, 202), 0u) == RGB(1, 2, 3), "COLORREF blend alpha zero preserves the under color");
    Require(BlendColorRefTruncate(RGB(1, 2, 3), RGB(200, 201, 202), 255u) == RGB(200, 201, 202), "COLORREF blend alpha 255 selects the over color");
    Require(BlendColorRefWeightedTruncate(RGB(10, 20, 30), RGB(110, 220, 230), 1, 2) == RGB(60, 120, 130),
            "weighted COLORREF blend preserves the existing arbitrary-denominator truncation policy");
    Require(BlendColorRefWeightedTruncate(RGB(10, 20, 30), RGB(110, 220, 230), -1, 10) == RGB(10, 20, 30),
            "weighted COLORREF blend clamps a negative overlay weight to zero");
    Require(BlendColorRefWeightedTruncate(RGB(10, 20, 30), RGB(110, 220, 230), 20, 10) == RGB(110, 220, 230),
            "weighted COLORREF blend clamps an oversized overlay weight to the denominator");
    Require(BlendColorRefWeightedTruncate(RGB(10, 20, 30), RGB(110, 220, 230), 1, 0) == RGB(10, 20, 30),
            "weighted COLORREF blend preserves the base color for an invalid denominator");
}

void TestStableVisualHash32Utf16V1GoldenValuesAndNamedEncodingPolicy()
{
    using Common::Colors::StableVisualHash32Utf16V1;

    Require(StableVisualHash32Utf16V1(L"") == 0x811C9DC5u, "visual hash v1 preserves the FNV-1a empty seed");
    Require(StableVisualHash32Utf16V1(L"abc") == 0x1A47E90Bu, "visual hash v1 preserves the ASCII golden value");
    Require(StableVisualHash32Utf16V1(L"\u00E9") == 0x6C0B6C44u, "visual hash v1 hashes a non-ASCII UTF-16 code unit as one value");
    Require(StableVisualHash32Utf16V1(L"\U0001F600") == 0xCB31C4B8u, "visual hash v1 hashes an astral character as its two UTF-16 surrogate code units");
}

void TestHsvColorRefNegativeHuePoliciesRemainExplicitAndDistinct()
{
    using namespace Common::Colors;

    Require(ColorRefFromHsvClampedNegativeHueToZero(-60.0f, 1.0f, 1.0f) == RGB(255, 0, 0), "majority viewer HSV policy clamps negative hue to red");
    Require(ColorRefFromHsvWrappedHue(-60.0f, 1.0f, 1.0f) == RGB(255, 0, 255), "ViewerWeb HSV policy wraps negative hue to magenta");
    Require(ColorRefFromHsvClampedNegativeHueToZero(120.0f, 2.0f, 2.0f) == RGB(0, 255, 0), "majority viewer HSV policy clamps saturation and value");
    Require(ColorRefFromHsvClampedNegativeHueToZero(240.0f, 1.0f, 1.0f) == RGB(0, 0, 255), "majority viewer HSV policy preserves primary blue");
}

void TestSharedLuminanceMathSeparatesLinearizedAndEncodedPolicies()
{
    using namespace Common::Colors;

    Require(std::abs(RelativeLuminanceFromArgb(0xFF000000u)) < 0.0000001, "relative luminance maps black to zero");
    Require(std::abs(RelativeLuminanceFromArgb(0xFFFFFFFFu) - 1.0) < 0.0000001, "relative luminance maps white to one");
    Require(std::abs(RelativeLuminanceFromArgb(0xFFFF0000u) - 0.2126) < 0.0000001, "relative luminance preserves the linear red coefficient");
    Require(std::abs(RelativeLuminanceFromColorRef(RGB(128, 128, 128)) - 0.2158605001) < 0.0000001, "relative luminance linearizes an encoded mid-gray");
    Require(std::abs(WeightedSrgbLuminanceWithoutLinearization(0.5, 0.5, 0.5) - 0.5) < 0.0000001,
            "encoded-channel luminance remains a separately named threshold policy");
    Require(std::abs(ContrastRatioFromRelativeLuminance(1.0, 0.0) - 21.0) < 0.0000001,
            "relative-luminance contrast ratio preserves the black/white golden value");
}

void TestD2dBlendClampsInterpolationAmount()
{
    using namespace RedSalamander::DxUi;

    const D2D1_COLOR_F from = D2D1::ColorF(0.1f, 0.2f, 0.3f, 0.4f);
    const D2D1_COLOR_F to   = D2D1::ColorF(0.8f, 0.7f, 0.6f, 0.5f);
    RequireColorNear(BlendColor(from, to, -1.0f), from, "D2D blend clamps negative interpolation to the source color");
    RequireColorNear(BlendColor(from, to, 2.0f), to, "D2D blend clamps interpolation above one to the destination color");
    RequireColorNear(BlendColor(from, to, 0.5f), D2D1::ColorF(0.45f, 0.45f, 0.45f, 0.45f), "D2D blend interpolates all four channels");
}

void TestWindowSizingDpiRoundingAndPopupGeometryGoldenValues()
{
    using namespace Common::WindowSizing;

    Require(DipToPixelRounded(96u, 8) == 8, "integer DIP conversion preserves values at 96 DPI");
    Require(DipToPixelRounded(120u, 8) == 10, "integer DIP conversion scales values at 120 DPI");
    Require(DipToPixelRounded(144u, 8) == 12, "integer DIP conversion scales values at 144 DPI");
    Require(DipToPixelRounded(192u, 8) == 16, "integer DIP conversion scales values at 192 DPI");
    Require(DipToPixelRounded(120u, -8) == -10, "integer DIP conversion preserves negative coordinates");
    Require(DipToPixelRounded(0.5f, 96u) == 1 && DipToPixelRounded(-0.5f, 96u) == -1, "fractional DIP conversion rounds half away from zero");
    Require(DipToPixelRounded((std::numeric_limits<float>::max)(), 192u) == (std::numeric_limits<int>::max)(),
            "fractional DIP conversion saturates positive overflow");
    Require(std::abs(PixelToDip(15.0f, 144.0f) - 10.0f) < 0.0001f, "pixel-to-DIP conversion preserves fractional monitor scaling");
    Require(std::abs(PixelToDip(15.0f, 0.0f) - 15.0f) < 0.0001f, "pixel-to-DIP conversion treats zero DPI as 96 DPI");

    Require(ComputeBoundedListPopupHeightPx(0u, 8u, 24, 10, 96u) == 34, "bounded popup height keeps one visible row for an empty list host");
    Require(ComputeBoundedListPopupHeightPx(100u, 8u, 24, 10, 192u) == 404, "bounded popup height caps visible rows before DPI scaling");

    const RECT work{-50, -40, 250, 160};
    const RECT clamped = ClampRectOriginToBounds(RECT{-200, -100, -100, 0}, work);
    Require(clamped.left == -50 && clamped.top == -40 && clamped.right == 50 && clamped.bottom == 60,
            "work-area clamping preserves size across negative monitor coordinates");
    const RECT oversized = ClampRectOriginToBounds(RECT{0, 0, 500, 300}, work);
    Require(oversized.left == -50 && oversized.top == -40 && oversized.right == 450 && oversized.bottom == 260,
            "oversized popup clamping pins the origin to the work-area origin");

    const RECT centered = CenterRectOnOwner(RECT{0, 0, 100, 50}, RECT{-300, -200, 300, 200});
    Require(centered.left == -50 && centered.top == -25 && centered.right == 50 && centered.bottom == 25,
            "owner centering preserves the window size across negative monitor coordinates");
    const RECT oddCentered = CenterRectOnOwner(RECT{10, 20, 111, 71}, RECT{0, 0, 400, 300});
    Require(oddCentered.left == 149 && oddCentered.top == 124 && oddCentered.right == 250 && oddCentered.bottom == 175,
            "owner centering preserves the existing integer midpoint rounding policy");

    const RECT verticalWork{-100, -50, 500, 450};
    Require(ResolveBoundedPopupTop(100, 120, 100, 4, verticalWork) == 124, "bounded popup placement prefers below when it fits");
    Require(ResolveBoundedPopupTop(400, 420, 100, 4, verticalWork) == 296, "bounded popup placement flips above near the bottom edge");
    Require(ResolveBoundedPopupTop(100, 120, 600, 4, verticalWork) == -50, "oversized popup placement clamps to the work-area top");
}

void TestViewerFileComboPopupInputAndHeightGoldenValues()
{
    using namespace RedSalamander::ViewerFileComboHost;

    Require(MessageMayOpenWindowComboPopup(WM_LBUTTONDOWN, 0u), "file combo opens for a primary-button press");
    Require(MessageMayOpenWindowComboPopup(WM_LBUTTONDBLCLK, 0u), "file combo opens for a primary-button double click");
    Require(MessageMayOpenWindowComboPopup(WM_SYSKEYDOWN, VK_DOWN) && MessageMayOpenWindowComboPopup(WM_SYSKEYDOWN, VK_UP),
            "file combo preserves Alt+Down and Alt+Up popup input");
    Require(MessageMayOpenWindowComboPopup(WM_KEYDOWN, VK_SPACE) && MessageMayOpenWindowComboPopup(WM_KEYDOWN, VK_RETURN) &&
                MessageMayOpenWindowComboPopup(WM_KEYDOWN, VK_F4) && MessageMayOpenWindowComboPopup(WM_KEYDOWN, VK_DOWN) &&
                MessageMayOpenWindowComboPopup(WM_KEYDOWN, VK_UP),
            "file combo preserves every keyboard popup input");
    Require(! MessageMayOpenWindowComboPopup(WM_KEYDOWN, VK_ESCAPE) && ! MessageMayOpenWindowComboPopup(WM_KEYUP, VK_F4) &&
                ! MessageMayOpenWindowComboPopup(WM_RBUTTONDOWN, 0u),
            "file combo rejects close keys, key-up, and secondary-button messages");
    Require(ComputeStandaloneComboPopupHeightPx(0u, 96u) == 34, "standalone file combo keeps one row plus chrome for an empty list");
    Require(ComputeStandaloneComboPopupHeightPx(100u, 192u) == 404, "standalone file combo caps its popup at eight visible rows before DPI scaling");
}

void TestViewerTitleBarThemeGoldenValues()
{
    using namespace RedSalamander::ViewerChrome;

    ViewerTheme theme{};
    theme.darkMode       = TRUE;
    theme.darkBase       = TRUE;
    theme.rainbowMode    = TRUE;
    theme.backgroundArgb = 0xFF101820u;
    theme.accentArgb     = 0xFF336699u;

    const COLORREF accent           = ResolveViewerThemeAccent(theme, L"title-golden");
    const TitleBarAttributes active = ResolveTitleBarAttributes(theme, true, L"title-golden");
    Require(active.useDarkMode == TRUE && active.borderColor == accent && active.captionColor == accent,
            "rainbow viewer title bar applies its deterministic accent while active");
    const uint32_t luma =
        (static_cast<uint32_t>(GetRValue(accent)) * 299u + static_cast<uint32_t>(GetGValue(accent)) * 587u + static_cast<uint32_t>(GetBValue(accent)) * 114u) /
        1000u;
    Require(active.textColor == (luma < 128u ? RGB(255, 255, 255) : RGB(0, 0, 0)), "rainbow viewer title bar preserves the integer contrast threshold");

    const TitleBarAttributes inactive = ResolveTitleBarAttributes(theme, false, L"title-golden");
    const COLORREF expectedInactive   = Common::Colors::BlendColorRefTruncate(accent, Common::Colors::ColorRefFromArgb(theme.backgroundArgb), 223u);
    Require(inactive.captionColor == expectedInactive && inactive.borderColor == expectedInactive,
            "inactive rainbow viewer title bar preserves the 223/255 background blend");

    theme.rainbowMode              = FALSE;
    const TitleBarAttributes plain = ResolveTitleBarAttributes(theme, true, L"ignored");
    Require(plain.useDarkMode == TRUE && plain.captionColor == kDwmColorDefault && plain.borderColor == kDwmColorDefault && plain.textColor == kDwmColorDefault,
            "non-rainbow viewer title bar resets custom DWM colors while retaining dark mode");

    theme.highContrast                    = TRUE;
    const TitleBarAttributes highContrast = ResolveTitleBarAttributes(theme, true, L"ignored");
    Require(highContrast.useDarkMode == FALSE && highContrast.captionColor == kDwmColorDefault,
            "high-contrast viewer title bar delegates all colors and dark-mode policy to the system");
}

void TestModifierCompositionCoversEveryCombination()
{
    using namespace RedSalamander::DxUi;

    for (UINT combination = 0u; combination < 256u; ++combination)
    {
        const UINT modifiers = ComposeModifierMask((combination & 0x01u) != 0u,
                                                   (combination & 0x02u) != 0u,
                                                   (combination & 0x04u) != 0u,
                                                   (combination & 0x08u) != 0u,
                                                   (combination & 0x10u) != 0u,
                                                   (combination & 0x20u) != 0u,
                                                   (combination & 0x40u) != 0u,
                                                   (combination & 0x80u) != 0u);
        const UINT expected  = ((combination & 0x01u) != 0u ? MK_SHIFT : 0u) | ((combination & 0x02u) != 0u ? MK_CONTROL : 0u) |
                               ((combination & 0x04u) != 0u ? kModifierAlt : 0u) | ((combination & 0x08u) != 0u ? MK_LBUTTON : 0u) |
                               ((combination & 0x10u) != 0u ? MK_RBUTTON : 0u) | ((combination & 0x20u) != 0u ? MK_MBUTTON : 0u) |
                               ((combination & 0x40u) != 0u ? MK_XBUTTON1 : 0u) | ((combination & 0x80u) != 0u ? MK_XBUTTON2 : 0u);
        Require(modifiers == expected, "modifier composition preserves every key/button combination");
        Require(ModifiersContainShift(modifiers) == ((combination & 0x01u) != 0u), "shift modifier decoding matches composition");
        Require(ModifiersContainCtrl(modifiers) == ((combination & 0x02u) != 0u), "control modifier decoding matches composition");
        Require(ModifiersContainAlt(modifiers) == ((combination & 0x04u) != 0u), "alt modifier decoding matches composition");
    }
}

void TestUtfConversionPoliciesRemainExplicitAndPreserveEmbeddedNulls()
{
    using namespace Common::Strings;

    const std::string malformedUtf8{"\xC3\x28", 2u};
    Require(! TryUtf16FromUtf8Strict(malformedUtf8).has_value(), "strict UTF-8 conversion rejects an invalid continuation byte");
    Require(Utf16FromUtf8ReplacingInvalid(malformedUtf8) == std::wstring{L"\uFFFD("},
            "replacement UTF-8 conversion retains displayable text after an invalid sequence");
    Require(! TryUtf16FromUtf8Strict(std::string{"\xF0\x9F\x98", 3u}).has_value(), "strict UTF-8 conversion rejects a truncated sequence");

    const std::string utf8WithNull{"A\0\xF0\x9F\x98\x80", 6u};
    const std::optional<std::wstring> utf16WithNull = TryUtf16FromUtf8Strict(utf8WithNull);
    Require(utf16WithNull.has_value() && utf16WithNull.value().size() == 4u && utf16WithNull.value()[0] == L'A' && utf16WithNull.value()[1] == L'\0' &&
                utf16WithNull.value()[2] == L'\xD83D' && utf16WithNull.value()[3] == L'\xDE00',
            "strict UTF-8 conversion preserves embedded NUL and supplementary characters");
    const std::optional<std::string> roundTrip = TryUtf8FromUtf16Strict(utf16WithNull.value());
    Require(roundTrip.has_value() && roundTrip.value() == utf8WithNull, "strict UTF conversion round-trips embedded NUL and supplementary characters");

    const std::wstring unpairedSurrogate{L'\xD800'};
    Require(! TryUtf8FromUtf16Strict(unpairedSurrogate).has_value(), "strict UTF-16 conversion rejects an unpaired surrogate");
    Require(Utf8FromUtf16ReplacingInvalid(unpairedSurrogate) == std::string{"\xEF\xBF\xBD", 3u},
            "replacement UTF-16 conversion emits U+FFFD for an unpaired surrogate");
    Require(TryUtf16FromUtf8Strict({}).has_value() && TryUtf16FromUtf8Strict({}).value().empty(),
            "strict UTF conversion distinguishes valid empty input from malformed input");
}

void TestOrdinalTrimAndEnvironmentPoliciesRemainExplicit()
{
    Require(OrdinalString::EqualsNoCase(L"FILE", L"file"), "ordinal identifier comparison is case-insensitive without locale folding");
    Require(! OrdinalString::EqualsNoCase(L"I", L"\u0131"), "ordinal identifier comparison does not apply Turkish dotless-I locale rules");
    Require(StringUtils::TrimWhitespace(L" \t\r\nvalue\u2003") == L"value", "shared whitespace trim preserves the existing iswspace character set");
    Require(StringUtils::TrimWhitespace(L"\u00A0value\u00A0") == L"value", "shared whitespace trim handles non-breaking-space boundaries");

    constexpr wchar_t kVariable[]              = L"REDSALAMANDER_LIGHTHOUSE_ENV_POLICY_TEST";
    const std::optional<std::wstring> previous = EnvironmentVariables::Read(kVariable);
    const auto restore = wil::scope_exit([&] { static_cast<void>(SetEnvironmentVariableW(kVariable, previous.has_value() ? previous->c_str() : nullptr)); });

    static_cast<void>(SetEnvironmentVariableW(kVariable, nullptr));
    Require(! EnvironmentVariables::Read(kVariable).has_value(), "environment reads distinguish a missing variable");
    static_cast<void>(SetEnvironmentVariableW(kVariable, L""));
    const std::optional<std::wstring> empty = EnvironmentVariables::Read(kVariable);
    Require(empty.has_value() && empty->empty(), "environment reads distinguish an explicitly empty variable");
    static_cast<void>(SetEnvironmentVariableW(kVariable, L"true"));
    Require(EnvironmentVariables::Read(kVariable).value_or(L"") == L"true", "environment reads preserve the complete value");
    Require(EnvironmentVariables::IsTruthyFlagSet(kVariable), "truthy environment flags accept the established t/y/1 prefixes");
    static_cast<void>(SetEnvironmentVariableW(kVariable, L"false"));
    Require(! EnvironmentVariables::IsTruthyFlagSet(kVariable), "truthy environment flags reject non-truthy prefixes");
}

void TestWindowsPathClassesKeepDriveRelativeAndFullyAbsolutePoliciesDistinct()
{
    using enum Common::Paths::WindowsPathClass;
    using Common::Paths::ClassifyWindowsPath;

    Require(ClassifyWindowsPath(L"folder\\leaf") == Relative, "relative Windows paths remain distinct");
    Require(ClassifyWindowsPath(L"\\folder\\leaf") == Rooted, "rooted current-drive paths are not fully absolute");
    Require(ClassifyWindowsPath(L"C:") == DriveRelative && ClassifyWindowsPath(L"C:leaf") == DriveRelative,
            "drive-qualified paths without a root separator remain drive-relative");
    Require(ClassifyWindowsPath(L"C:\\leaf") == DriveAbsolute && ClassifyWindowsPath(L"d:/leaf") == DriveAbsolute,
            "drive-absolute classification accepts both Windows separator variants");
    Require(ClassifyWindowsPath(L"\\\\server\\share\\leaf") == Unc && ClassifyWindowsPath(L"//server/share/leaf") == Unc,
            "UNC classification accepts slash variants without treating them as drive paths");
    Require(ClassifyWindowsPath(L"\\\\?\\C:\\leaf") == ExtendedDriveAbsolute && ClassifyWindowsPath(L"//?/C:/leaf") == ExtendedDriveAbsolute,
            "extended drive paths remain a named class");
    Require(ClassifyWindowsPath(L"\\\\?\\UNC\\server\\share\\leaf") == ExtendedUnc, "extended UNC paths remain distinct from ordinary UNC paths");
    Require(ClassifyWindowsPath(L"\\\\?\\Volume{guid}\\leaf") == ExtendedOther,
            "non-drive extended namespaces are not silently treated as fully absolute filesystem paths");
    Require(ClassifyWindowsPath(L"\\\\.\\PhysicalDrive0") == Device, "device namespace paths remain explicitly classified");

    Require(Common::Paths::IsDriveQualifiedWindowsPath(L"C:") && ! Common::Paths::IsDriveAbsoluteWindowsPath(L"C:"),
            "drive-qualified and drive-absolute predicates remain separate");
    Require(Common::Paths::IsFullyAbsoluteWindowsPath(L"C:\\leaf") && Common::Paths::IsFullyAbsoluteWindowsPath(L"\\\\server\\share") &&
                Common::Paths::IsFullyAbsoluteWindowsPath(L"\\\\?\\UNC\\server\\share\\leaf"),
            "fully absolute policy accepts rooted drives and complete UNC server/share roots");
    Require(! Common::Paths::IsFullyAbsoluteWindowsPath(L"C:leaf") && ! Common::Paths::IsFullyAbsoluteWindowsPath(L"\\\\server") &&
                ! Common::Paths::IsFullyAbsoluteWindowsPath(L"\\\\.\\PhysicalDrive0"),
            "fully absolute policy rejects drive-relative, incomplete UNC, and device paths");
    Require(ClassifyWindowsPath(L"C:\\dir\\..\\CON. ") == DriveAbsolute, "classification remains normalization-free and does not weaken leaf validation");
    Require(Common::Paths::ToExtendedWin32Path(L"C:\\leaf") == LR"(\\?\C:\leaf)", "normalization-free extended conversion prefixes drive-absolute paths");
    Require(Common::Paths::ToExtendedWin32Path(L"\\\\server\\share\\leaf") == LR"(\\?\UNC\server\share\leaf)",
            "normalization-free extended conversion prefixes UNC paths");
    Require(Common::Paths::ToExtendedWin32Path(L"C:leaf") == L"C:leaf" &&
                Common::Paths::ToExtendedWin32Path(L"\\\\.\\PhysicalDrive0") == L"\\\\.\\PhysicalDrive0",
            "extended conversion leaves drive-relative and device paths unchanged");
}

void TestBinaryThroughputGrammarPreservesBoundaryPoliciesAndRoundTrips()
{
    using Common::Parsing::FormatBinaryThroughputText;
    using Common::Parsing::ThroughputBoundaryWhitespacePolicy;
    using Common::Parsing::TryParseBinaryThroughputText;

    constexpr ThroughputBoundaryWhitespacePolicy ascii = ThroughputBoundaryWhitespacePolicy::AsciiWhitespace;
    constexpr ThroughputBoundaryWhitespacePolicy c0    = ThroughputBoundaryWhitespacePolicy::ControlCharactersThroughSpace;

    uint64_t parsed = 99u;
    Require(TryParseBinaryThroughputText(L"", ascii, parsed) && parsed == 0u, "empty throughput text preserves the unlimited-zero policy");
    Require(TryParseBinaryThroughputText(L" \t1.5 MiB/s\r\n", ascii, parsed) && parsed == 1572864u,
            "throughput parsing accepts fractional values, ASCII boundary whitespace, and the per-second suffix");
    Require(TryParseBinaryThroughputText(L"1,5 mIB/S", ascii, parsed) && parsed == 1572864u,
            "throughput parsing accepts the locale-independent comma decimal and case-insensitive binary alias");
    Require(TryParseBinaryThroughputText(L".5", ascii, parsed) && parsed == 512u, "bare fractional throughput values default to KiB/s");
    Require(TryParseBinaryThroughputText(L"0.0005 KiB", ascii, parsed) && parsed == 1u, "fractional throughput rounds to the nearest whole byte");
    Require(TryParseBinaryThroughputText(L"2 b", ascii, parsed) && parsed == 2u && TryParseBinaryThroughputText(L"2 K", ascii, parsed) && parsed == 2048u &&
                TryParseBinaryThroughputText(L"2 MB", ascii, parsed) && parsed == 2097152u && TryParseBinaryThroughputText(L"2 g", ascii, parsed) &&
                parsed == 2147483648u && TryParseBinaryThroughputText(L"2 TiB", ascii, parsed) && parsed == 2199023255552u &&
                TryParseBinaryThroughputText(L"2 p", ascii, parsed) && parsed == 2251799813685248u,
            "throughput parsing preserves every established byte-through-PiB alias");
    Require(TryParseBinaryThroughputText(L"999999999999999999999999999 PiB", ascii, parsed) && parsed == (std::numeric_limits<uint64_t>::max)(),
            "throughput parsing saturates positive overflow");

    parsed = 99u;
    Require(! TryParseBinaryThroughputText(L"1 XB", ascii, parsed) && parsed == 0u, "unknown throughput units fail and clear the output");
    Require(! TryParseBinaryThroughputText(L"1.2.3 MiB", ascii, parsed), "multiple decimal separators remain invalid");
    Require(! TryParseBinaryThroughputText(L"-1 MiB", ascii, parsed), "signed throughput values remain invalid");

    std::wstring controlWrapped = L"1 MiB";
    controlWrapped.insert(controlWrapped.begin(), static_cast<wchar_t>(0x001F));
    controlWrapped.push_back(static_cast<wchar_t>(0x001F));
    Require(TryParseBinaryThroughputText(controlWrapped, c0, parsed) && parsed == 1048576u,
            "Preferences preserves its historical C0-through-space boundary trimming");
    Require(! TryParseBinaryThroughputText(controlWrapped, ascii, parsed), "the popup preserves its narrower six-character ASCII whitespace boundary");

    constexpr std::array<uint64_t, 7u> roundTripValues{{0u, 1u, 1024u, 1048576u, 1073741824u, 1048577u, (std::numeric_limits<uint64_t>::max)()}};
    for (const uint64_t value : roundTripValues)
    {
        const std::wstring formatted = FormatBinaryThroughputText(value);
        Require(TryParseBinaryThroughputText(formatted, c0, parsed) && parsed == value,
                "the File Operations throughput formatter round-trips through the shared parser");
    }
}

void TestCompactFileMetadataFormattingPreservesTimeAndAttributePolicies()
{
    using Common::FileMetadata::DisplayProfile;
    using Common::FileMetadata::FormatDisplayFields;
    using Common::FileMetadata::NormalizedMetadata;

    const auto empty = FormatDisplayFields({}, DisplayProfile::CompactDetails);
    Require(empty.localTime.empty() && empty.attributes == L"-", "missing time and zero attributes preserve the compact empty/dash representation");

    const auto attributes =
        FormatDisplayFields({.fileAttributes = FILE_ATTRIBUTE_READONLY | FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM | FILE_ATTRIBUTE_ARCHIVE |
                                               FILE_ATTRIBUTE_COMPRESSED | FILE_ATTRIBUTE_ENCRYPTED | FILE_ATTRIBUTE_TEMPORARY | FILE_ATTRIBUTE_OFFLINE |
                                               FILE_ATTRIBUTE_REPARSE_POINT | FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_SPARSE_FILE},
                            DisplayProfile::CompactDetails);
    Require(attributes.attributes == L"RHSACETOP", "compact metadata preserves RHSACETOP ordering and intentionally omits directory and sparse flags");
    const auto sparseCompressed =
        FormatDisplayFields({.fileAttributes = FILE_ATTRIBUTE_SPARSE_FILE | FILE_ATTRIBUTE_COMPRESSED}, DisplayProfile::CompactDetails);
    Require(sparseCompressed.attributes == L"C", "compressed remains visible while sparse remains outside the compact attribute profile");

    Require(FormatDisplayFields({.lastWriteTime100nsSince1601 = -1}, DisplayProfile::CompactDetails).localTime.empty(),
            "negative metadata times remain unknown");
    Require(FormatDisplayFields({.lastWriteTime100nsSince1601 = (std::numeric_limits<int64_t>::max)()}, DisplayProfile::CompactDetails).localTime.empty(),
            "invalid metadata times fail closed to an empty display field");

    SYSTEMTIME utc{};
    utc.wYear   = 2024u;
    utc.wMonth  = 1u;
    utc.wDay    = 15u;
    utc.wHour   = 12u;
    utc.wMinute = 34u;
    utc.wSecond = 56u;
    FILETIME utcFileTime{};
    Require(SystemTimeToFileTime(&utc, &utcFileTime) != FALSE, "metadata test creates a valid UTC FILETIME");
    ULARGE_INTEGER ticks{};
    ticks.LowPart  = utcFileTime.dwLowDateTime;
    ticks.HighPart = utcFileTime.dwHighDateTime;

    FILETIME localFileTime{};
    SYSTEMTIME localSystemTime{};
    Require(FileTimeToLocalFileTime(&utcFileTime, &localFileTime) != FALSE && FileTimeToSystemTime(&localFileTime, &localSystemTime) != FALSE,
            "metadata test resolves the current Windows time-zone conversion");
    const std::wstring expectedLocal = std::format(L"{:04d}-{:02d}-{:02d} {:02d}:{:02d}",
                                                   localSystemTime.wYear,
                                                   localSystemTime.wMonth,
                                                   localSystemTime.wDay,
                                                   localSystemTime.wHour,
                                                   localSystemTime.wMinute);
    Require(FormatDisplayFields({.lastWriteTime100nsSince1601 = static_cast<int64_t>(ticks.QuadPart)}, DisplayProfile::CompactDetails).localTime ==
                expectedLocal,
            "compact metadata uses Windows local time, fixed date ordering, and minute precision independent of display locale");

    SYSTEMTIME preUnixUtc{};
    preUnixUtc.wYear  = 1969u;
    preUnixUtc.wMonth = 12u;
    preUnixUtc.wDay   = 31u;
    preUnixUtc.wHour  = 12u;
    FILETIME preUnixFileTime{};
    Require(SystemTimeToFileTime(&preUnixUtc, &preUnixFileTime) != FALSE, "metadata test creates a valid pre-Unix-epoch FILETIME");
    ticks.LowPart  = preUnixFileTime.dwLowDateTime;
    ticks.HighPart = preUnixFileTime.dwHighDateTime;
    Require(! FormatDisplayFields({.lastWriteTime100nsSince1601 = static_cast<int64_t>(ticks.QuadPart)}, DisplayProfile::CompactDetails).localTime.empty(),
            "valid pre-Unix-epoch Windows times remain displayable");

    Require(FormatBytesCompact(0u) == L"0 B", "the existing shared size formatter preserves exact zero-size display");
}

void TestPercentEncodingAndHandleIoPoliciesCoverForwardProgressAndAtomicSiblings()
{
    using Common::Uri::SlashPolicy;

    Require(Common::Uri::PercentEncodeBytes("AZaz09-._~") == "AZaz09-._~", "RFC3986 unreserved bytes remain literal");
    Require(Common::Uri::PercentEncodeBytes(" /?#%") == "%20%2F%3F%23%25", "reserved and unsafe bytes use uppercase percent escapes");
    Require(Common::Uri::PercentEncodeBytes("a/b c", SlashPolicy::Preserve) == "a/b%20c", "slash preservation is an explicit protocol policy");
    Require(Common::Uri::PercentEncodeBytes(std::string_view("\xC3\xA9", 2u)) == "%C3%A9", "UTF-8 is encoded byte-by-byte");
    std::wstring encodedWide;
    Require(Common::Uri::TryPercentEncodeUtf8ToWide(L"\U0001F600/x", SlashPolicy::Preserve, encodedWide) && encodedWide == L"%F0%9F%98%80/x",
            "wide protocol input is strictly converted to UTF-8 before byte encoding");
    const wchar_t unpaired[] = {static_cast<wchar_t>(0xD800), L'\0'};
    Require(! Common::Uri::TryPercentEncodeUtf8ToWide(unpaired, SlashPolicy::Encode, encodedWide),
            "strict protocol encoding rejects an unpaired UTF-16 surrogate");

    const std::array<std::byte, 5u> payload{};
    unsigned int partialCalls = 0u;
    const HRESULT partialHr =
        Common::HandleIo::Detail::TransferAll<const std::byte>(payload,
                                                               ERROR_WRITE_FAULT,
                                                               [&partialCalls](const std::byte* data, DWORD requested, DWORD& transferred) noexcept
    {
        static_cast<void>(data);
        ++partialCalls;
        transferred = (std::min)(requested, DWORD{2u});
        return S_OK;
    });
    Require(partialHr == S_OK && partialCalls == 3u, "total transfer retries partial forward progress until the buffer is complete");
    const HRESULT zeroProgressHr =
        Common::HandleIo::Detail::TransferAll<const std::byte>(payload,
                                                               ERROR_WRITE_FAULT,
                                                               [](const std::byte* data, DWORD requested, DWORD& transferred) noexcept
    {
        static_cast<void>(data);
        static_cast<void>(requested);
        transferred = 0u;
        return S_OK;
    });
    Require(zeroProgressHr == HRESULT_FROM_WIN32(ERROR_WRITE_FAULT), "successful zero-byte writes fail instead of spinning forever");
    const HRESULT errorHr = Common::HandleIo::Detail::TransferAll<const std::byte>(payload,
                                                                                   ERROR_WRITE_FAULT,
                                                                                   [](const std::byte* data, DWORD requested, DWORD& transferred) noexcept
    {
        static_cast<void>(data);
        static_cast<void>(requested);
        transferred = 0u;
        return HRESULT_FROM_WIN32(ERROR_DISK_FULL);
    });
    Require(errorHr == HRESULT_FROM_WIN32(ERROR_DISK_FULL), "total transfer preserves the underlying I/O error");

    const std::wstring sibling = GetDxUiTestArtifactPath(L"lighthouse-target.bin").wstring();
    const Common::Paths::UniqueSiblingFileOptions options{.prefix = L"lighthouse-", .suffix = L".tmp", .desiredAccess = GENERIC_READ | GENERIC_WRITE};
    std::wstring firstPath;
    std::wstring secondPath;
    wil::unique_hfile firstFile;
    wil::unique_hfile secondFile;
    Require(SUCCEEDED(Common::Paths::CreateUniqueSiblingFile(sibling, options, firstPath, firstFile)) && firstFile,
            "atomic sibling creation returns ownership of the CREATE_NEW reservation");
    Require(SUCCEEDED(Common::Paths::CreateUniqueSiblingFile(sibling, options, secondPath, secondFile)) && secondFile && secondPath != firstPath,
            "independent atomic sibling reservations cannot publish the same path");
    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        firstFile.reset();
        secondFile.reset();
        if (! firstPath.empty())
        {
            static_cast<void>(DeleteFileW(firstPath.c_str()));
        }
        if (! secondPath.empty())
        {
            static_cast<void>(DeleteFileW(secondPath.c_str()));
        }
    });

    constexpr std::array<std::byte, 5u> bytes{{std::byte{'R'}, std::byte{'S'}, std::byte{'I'}, std::byte{'O'}, std::byte{'!'}}};
    Require(Common::HandleIo::WriteAll(firstFile.get(), bytes) == S_OK, "shared handle writer writes the complete payload");
    uint64_t fileSize = 0u;
    Require(Common::HandleIo::GetFileSizeBounded(firstFile.get(), bytes.size(), fileSize) == S_OK && fileSize == bytes.size(),
            "bounded handle size accepts a file at the configured maximum");
    Require(Common::HandleIo::GetFileSizeBounded(firstFile.get(), bytes.size() - 1u, fileSize) == HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE),
            "bounded handle size rejects oversized files at the domain boundary");
    Require(Common::HandleIo::Rewind(firstFile.get()) == S_OK, "shared handle rewind restores the beginning position");
    std::array<std::byte, bytes.size()> readBack{};
    Require(Common::HandleIo::ReadExact(firstFile.get(), readBack) == S_OK && readBack == bytes, "shared exact reader returns the entire payload after rewind");

    std::wstring stalePath = L"stale";
    wil::unique_hfile staleFile;
    const Common::Paths::UniqueSiblingFileOptions invalidOptions{.prefix = L"bad\\prefix", .suffix = L".tmp"};
    Require(Common::Paths::CreateUniqueSiblingFile(sibling, invalidOptions, stalePath, staleFile) == E_INVALIDARG && stalePath.empty() && ! staleFile,
            "failed atomic sibling validation clears unpublished path and handle outputs");
}

void TestCloudPaginationGuardBoundsProgressAndCancellation()
{
    using Common::Paging::Limits;
    using Common::Paging::Utf8ContinuationGuard;

    bool cancel                  = false;
    const auto cancellationProbe = [](void* cookie) noexcept -> HRESULT
    { return *static_cast<const bool*>(cookie) ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK; };

    Limits limits{
        .maxPages           = 2u,
        .maxItems           = 3u,
        .maxBytes           = 12u,
        .maxTokenChars      = 8u,
        .deadlineTickMs     = 110u,
        .cancellationProbe  = cancellationProbe,
        .cancellationCookie = &cancel,
    };
    Utf8ContinuationGuard guard(limits);
    Require(guard.BeginFirstPage(100u) == S_OK, "pagination guard admits the bounded first page");
    Require(guard.CompletePage(2u, 5u, true, "next", 101u) == S_OK, "pagination guard records page items/bytes and a non-empty continuation");
    Require(guard.BeginContinuation("next", 102u) == S_OK, "pagination guard admits a new continuation once");
    Require(guard.CompletePage(1u, 7u, false, {}, 103u) == S_OK && guard.PageCount() == 2u && guard.ItemCount() == 3u && guard.ByteCount() == 12u,
            "pagination guard preserves exact bounded totals");

    Utf8ContinuationGuard repeated(limits);
    Require(repeated.BeginFirstPage(100u) == S_OK && repeated.CompletePage(0u, 0u, true, "same", 101u) == S_OK &&
                repeated.BeginContinuation("same", 102u) == S_OK && repeated.CompletePage(0u, 0u, true, "same", 103u) == S_OK &&
                repeated.BeginContinuation("same", 104u) == HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
            "pagination guard rejects a repeated continuation before another request");

    Utf8ContinuationGuard empty(limits);
    Require(empty.BeginFirstPage(100u) == S_OK && empty.CompletePage(0u, 0u, true, {}, 101u) == HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
            "pagination guard rejects an empty token when the provider says another page exists");

    Limits tiny         = limits;
    tiny.maxPages       = 1u;
    tiny.deadlineTickMs = 0u;
    Utf8ContinuationGuard pageCap(tiny);
    Require(pageCap.BeginFirstPage(100u) == S_OK && pageCap.CompletePage(0u, 0u, true, "next", 101u) == S_OK &&
                pageCap.BeginContinuation("next", 102u) == HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE),
            "pagination guard enforces the page ceiling");

    Utf8ContinuationGuard deadline(limits);
    Require(deadline.BeginFirstPage(110u) == HRESULT_FROM_WIN32(ERROR_TIMEOUT), "pagination guard enforces the operation deadline");

    cancel                         = true;
    Limits cancelledLimits         = limits;
    cancelledLimits.deadlineTickMs = 0u;
    Utf8ContinuationGuard cancelled(cancelledLimits);
    Require(cancelled.BeginFirstPage(100u) == HRESULT_FROM_WIN32(ERROR_CANCELLED), "pagination guard propagates cancellation at request boundaries");
}

void TestYyjsonMemberPoliciesDistinguishPresenceTypeRangeAndCoercion()
{
    using namespace Common::Json;

    std::string json =
        R"json({"text":"value","nullValue":null,"wrong":true,"signed":-7,"unsigned":9,"tooLarge":18446744073709551615,"numericString":"42","unknownFuture":{"x":1}})json";
    UniqueDocument document{yyjson_read(json.data(), json.size(), 0)};
    Require(document != nullptr, "shared yyjson ownership accepts a valid document");
    const yyjson_val* root = yyjson_doc_get_root(document.get());

    const auto text = GetStringMember(root, "text", MemberRequirement::Required);
    Require(text.HasValue() && text.value == "value", "strict string member access returns the borrowed value with its explicit length");
    Require(GetStringMember(root, "missing", MemberRequirement::Optional).status == MemberStatus::MissingOptional,
            "optional member access distinguishes a missing optional member");
    Require(GetStringMember(root, "missing", MemberRequirement::Required).status == MemberStatus::MissingRequired,
            "required member access distinguishes a missing required member");
    Require(GetStringMember(root, "nullValue", MemberRequirement::Optional).status == MemberStatus::Null, "member access distinguishes JSON null from missing");
    Require(GetStringMember(root, "wrong", MemberRequirement::Required).status == MemberStatus::WrongType,
            "strict string member access rejects the wrong JSON type");

    const auto signedValue = GetInt64Member(root, "signed", MemberRequirement::Required);
    Require(signedValue.HasValue() && signedValue.value == -7, "strict signed member access accepts a signed integer");
    Require(GetInt64Member(root, "tooLarge", MemberRequirement::Required).status == MemberStatus::OutOfRange, "signed member access reports unsigned overflow");
    Require(GetInt64Member(root, "numericString", MemberRequirement::Required).status == MemberStatus::WrongType,
            "strict integer access rejects numeric strings");
    const auto coerced = GetInt64Member(root, "numericString", MemberRequirement::Required, NumericStringPolicy::Allow);
    Require(coerced.HasValue() && coerced.value == 42, "numeric-string coercion occurs only when requested at the call site");
    Require(GetUInt64Member(root, "signed", MemberRequirement::Required).status == MemberStatus::OutOfRange,
            "unsigned member access reports negative values as out of range");
    Require(
        GetUInt64Member(root, "unsigned", MemberRequirement::Required, NumericStringPolicy::Reject, UnsignedIntegerPolicy::RequireUnsignedStorage).HasValue(),
        "unsigned-storage member access accepts a JSON unsigned integer");
    Require(GetUInt64Member(root, "signed", MemberRequirement::Required, NumericStringPolicy::Reject, UnsignedIntegerPolicy::RequireUnsignedStorage).status ==
                MemberStatus::WrongType,
            "unsigned-storage member access rejects signed storage even before range conversion");
    Require(GetBoolMember(root, "wrong", MemberRequirement::Required).HasValue(), "strict bool member access accepts only a JSON boolean");
    Require(GetBoolMember(root, "unsigned", MemberRequirement::Required).status == MemberStatus::WrongType,
            "strict bool member access rejects integer coercion by default");
    const auto coercedBool = GetBoolMember(root, "unsigned", MemberRequirement::Required, BooleanIntegerPolicy::AllowZeroAndNonzero);
    Require(coercedBool.HasValue() && coercedBool.value, "integer-to-bool coercion occurs only when requested at the call site");
}

void TestViewerThemePaletteDerivesDarkControlChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF101214u;
    viewerTheme.textArgb                   = 0xFFE7E9EDu;
    viewerTheme.selectionBackgroundArgb    = 0xFF2A6DB2u;
    viewerTheme.selectionTextArgb          = 0xFFF8FAFCu;
    viewerTheme.accentArgb                 = 0xFF3A82D6u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = TRUE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette palette = MakeThemePaletteFromViewerTheme(viewerTheme);
    Require(palette.dark, "dark viewer theme keeps the palette in dark mode");
    Require(palette.rainbowMode, "dark viewer theme preserves rainbow mode");
    Require(! palette.highContrast, "dark viewer theme keeps high contrast disabled when requested");
    RequireColorNear(
        palette.windowBackground, D2D1::ColorF(0x10 / 255.0f, 0x12 / 255.0f, 0x14 / 255.0f, 1.0f), "dark viewer theme keeps the requested window background");
    RequireColorNear(palette.text, D2D1::ColorF(0xE7 / 255.0f, 0xE9 / 255.0f, 0xED / 255.0f, 1.0f), "dark viewer theme keeps the requested text color");
    RequireColorNear(palette.surfaceBackground,
                     BlendForTest(palette.windowBackground, palette.selectionFill, 0.05f),
                     "dark viewer theme derives surface background from the palette selection fill");
    RequireColorDifferent(palette.surfaceBackground, palette.accent, "dark viewer theme surface background follows selection fill instead of raw accent");
    RequireColorNear(palette.headerBackground,
                     BlendForTest(palette.surfaceBackground, palette.selectionFill, 0.12f),
                     "dark viewer theme derives header background from the palette selection fill");
    RequireColorDifferent(palette.headerBackground, palette.accent, "dark viewer theme header background follows selection fill instead of raw accent");
    RequireColorNear(palette.headerHovered,
                     BlendForTest(palette.headerBackground, D2D1::ColorF(palette.hoverFill.r, palette.hoverFill.g, palette.hoverFill.b, 1.0f), 0.22f),
                     "dark viewer theme derives hovered header chrome from the palette hover overlay");
    RequireColorNear(palette.headerPressed,
                     BlendForTest(palette.headerBackground, D2D1::ColorF(palette.pressedFill.r, palette.pressedFill.g, palette.pressedFill.b, 1.0f), 0.32f),
                     "dark viewer theme derives pressed header chrome from the palette pressed overlay");
    RequireColorDifferent(palette.border, palette.windowBackground, "dark viewer theme derives border chrome from the live colors");
    RequireColorDifferent(palette.gridLine, palette.surfaceBackground, "dark viewer theme derives grid lines from the live colors");
    RequireColorDifferent(palette.subduedText, palette.text, "dark viewer theme derives subdued text from the live colors");
    RequireColorDifferent(palette.disabledText, palette.text, "dark viewer theme derives disabled text from the live colors");
    RequireColorNear(palette.focusStroke,
                     BlendForTest(palette.selectionFill, palette.selectionText, 0.10f),
                     "dark viewer theme derives focus stroke chrome from selection fill and selection text");
    RequireColorDifferent(palette.focusStroke, palette.accent, "dark viewer theme focus stroke follows selection fill instead of raw accent");
    RequireColorNear(palette.hoverFill,
                     D2D1::ColorF(palette.focusStroke.r, palette.focusStroke.g, palette.focusStroke.b, 0.16f),
                     "dark viewer theme derives hover fill chrome from the palette focus stroke");
    RequireColorNear(palette.pressedFill,
                     D2D1::ColorF(palette.focusStroke.r, palette.focusStroke.g, palette.focusStroke.b, 0.24f),
                     "dark viewer theme derives pressed fill chrome from the palette focus stroke");
    RequireColorNear(palette.buttonFill,
                     BlendForTest(palette.surfaceBackground, palette.selectionFill, 0.04f),
                     "dark viewer theme derives button fill from the palette selection fill");
    RequireColorDifferent(palette.buttonFill, palette.accent, "dark viewer theme button fill follows selection fill instead of raw accent");
    RequireColorNear(palette.buttonBorder,
                     BlendForTest(palette.border, palette.focusStroke, 0.12f),
                     "dark viewer theme derives button border chrome from the palette focus stroke");
    RequireColorNear(palette.buttonHotFill,
                     BlendForTest(palette.buttonFill, palette.focusStroke, 0.10f),
                     "dark viewer theme derives button hot fill chrome from the palette focus stroke");
    RequireColorNear(palette.buttonPressedFill,
                     BlendForTest(palette.buttonFill, palette.focusStroke, 0.18f),
                     "dark viewer theme derives button pressed fill chrome from the palette focus stroke");
    RequireColorNear(palette.inputBorder,
                     BlendForTest(palette.border, palette.focusStroke, 0.18f),
                     "dark viewer theme derives input border chrome from the palette focus stroke");
    RequireColorDifferent(palette.inputFill, palette.windowBackground, "dark viewer theme derives input fill from the live colors");
    RequireColorDifferent(palette.tooltipBackground, palette.windowBackground, "dark viewer theme derives tooltip background from the live colors");
    RequireColorDifferent(palette.tooltipText, palette.tooltipBackground, "dark viewer theme keeps tooltip text legible against the tooltip background");
    RequireColorNear(palette.toggleKnobFill,
                     ChooseContrastingTextColorForTest(BlendForTest(palette.inputFill, palette.border, 0.18f)),
                     "dark viewer theme derives neutral toggle knob chrome from the live neutral palette");
    RequireColorNear(palette.toggleKnobCheckedFill,
                     D2D1::ColorF(0xF8 / 255.0f, 0xFA / 255.0f, 0xFC / 255.0f, 1.0f),
                     "dark viewer theme keeps selection text for checked toggle knob chrome");
    RequireColorNear(
        palette.selectionText, D2D1::ColorF(0xF8 / 255.0f, 0xFA / 255.0f, 0xFC / 255.0f, 1.0f), "dark viewer theme keeps the requested selection text chrome");
    Require(palette.scrollbarThumbHot.a > palette.scrollbarThumb.a, "dark viewer theme keeps the hot scrollbar thumb stronger than the idle thumb");
    Require(palette.pressedFill.a > palette.hoverFill.a, "dark viewer theme keeps pressed fill stronger than hover fill");
}

void TestViewerThemePaletteDerivesLightControlChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFFF7F4EEu;
    viewerTheme.textArgb                   = 0xFF24292Fu;
    viewerTheme.selectionBackgroundArgb    = 0xFFD9E8FFu;
    viewerTheme.selectionTextArgb          = 0xFF0E223Bu;
    viewerTheme.accentArgb                 = 0xFFB85C1Eu;
    viewerTheme.alertErrorBackgroundArgb   = 0xFFFFE2E0u;
    viewerTheme.alertErrorTextArgb         = 0xFF7D2018u;
    viewerTheme.alertWarningBackgroundArgb = 0xFFFFF1D2u;
    viewerTheme.alertWarningTextArgb       = 0xFF815100u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFFE2F0FFu;
    viewerTheme.alertInfoTextArgb          = 0xFF194A78u;
    viewerTheme.darkMode                   = FALSE;
    viewerTheme.highContrast               = TRUE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = FALSE;

    const ThemePalette palette = MakeThemePaletteFromViewerTheme(viewerTheme);
    Require(! palette.dark, "light viewer theme keeps the palette in light mode");
    Require(! palette.rainbowMode, "light viewer theme keeps rainbow mode disabled when requested");
    Require(palette.highContrast, "light viewer theme preserves high contrast");
    RequireColorNear(
        palette.windowBackground, D2D1::ColorF(0xF7 / 255.0f, 0xF4 / 255.0f, 0xEE / 255.0f, 1.0f), "light viewer theme keeps the requested window background");
    RequireColorNear(palette.text, D2D1::ColorF(0x24 / 255.0f, 0x29 / 255.0f, 0x2F / 255.0f, 1.0f), "light viewer theme keeps the requested text color");
    RequireColorNear(palette.surfaceBackground,
                     BlendForTest(palette.windowBackground, palette.selectionFill, 0.02f),
                     "light viewer theme derives surface background from the palette selection fill");
    RequireColorDifferent(palette.surfaceBackground, palette.accent, "light viewer theme surface background follows selection fill instead of raw accent");
    RequireColorNear(palette.headerBackground,
                     BlendForTest(palette.surfaceBackground, palette.selectionFill, 0.05f),
                     "light viewer theme derives header background from the palette selection fill");
    RequireColorDifferent(palette.headerBackground, palette.accent, "light viewer theme header background follows selection fill instead of raw accent");
    RequireColorNear(palette.headerHovered,
                     BlendForTest(palette.headerBackground, D2D1::ColorF(palette.hoverFill.r, palette.hoverFill.g, palette.hoverFill.b, 1.0f), 0.10f),
                     "light viewer theme derives hovered header chrome from the palette hover overlay");
    RequireColorNear(palette.headerPressed,
                     BlendForTest(palette.headerBackground, D2D1::ColorF(palette.pressedFill.r, palette.pressedFill.g, palette.pressedFill.b, 1.0f), 0.16f),
                     "light viewer theme derives pressed header chrome from the palette pressed overlay");
    RequireColorNear(palette.focusStroke,
                     BlendForTest(palette.selectionFill, palette.selectionText, 0.04f),
                     "light viewer theme derives focus stroke chrome from selection fill and selection text");
    RequireColorDifferent(palette.focusStroke, palette.accent, "light viewer theme focus stroke follows selection fill instead of raw accent");
    RequireColorNear(palette.hoverFill,
                     D2D1::ColorF(palette.focusStroke.r, palette.focusStroke.g, palette.focusStroke.b, 0.30f),
                     "light viewer theme derives hover fill chrome from the palette focus stroke in high contrast");
    RequireColorNear(palette.pressedFill,
                     D2D1::ColorF(palette.focusStroke.r, palette.focusStroke.g, palette.focusStroke.b, 0.38f),
                     "light viewer theme derives pressed fill chrome from the palette focus stroke in high contrast");
    RequireColorNear(palette.buttonFill,
                     BlendForTest(palette.surfaceBackground, palette.selectionFill, 0.02f),
                     "light viewer theme derives button fill from the palette selection fill");
    RequireColorDifferent(palette.buttonFill, palette.accent, "light viewer theme button fill follows selection fill instead of raw accent");
    RequireColorNear(palette.buttonBorder,
                     BlendForTest(palette.border, palette.focusStroke, 0.06f),
                     "light viewer theme derives button border chrome from the palette focus stroke");
    RequireColorNear(palette.buttonHotFill,
                     BlendForTest(palette.buttonFill, palette.focusStroke, 0.05f),
                     "light viewer theme derives button hot fill chrome from the palette focus stroke");
    RequireColorNear(palette.buttonPressedFill,
                     BlendForTest(palette.buttonFill, palette.focusStroke, 0.10f),
                     "light viewer theme derives button pressed fill chrome from the palette focus stroke");
    RequireColorNear(palette.inputBorder,
                     BlendForTest(palette.border, palette.focusStroke, 0.08f),
                     "light viewer theme derives input border chrome from the palette focus stroke");
    RequireColorDifferent(palette.inputFill, palette.windowBackground, "light viewer theme derives input fill from the live colors");
    RequireColorDifferent(palette.tooltipBackground, palette.windowBackground, "light viewer theme derives tooltip background from the live colors");
    RequireColorDifferent(palette.tooltipText, palette.tooltipBackground, "light viewer theme keeps tooltip text legible against the tooltip background");
    RequireColorNear(palette.toggleKnobFill,
                     ChooseContrastingTextColorForTest(BlendForTest(palette.inputFill, palette.border, 0.08f)),
                     "light viewer theme derives neutral toggle knob chrome from the live neutral palette");
    RequireColorNear(palette.toggleKnobCheckedFill,
                     D2D1::ColorF(0x0E / 255.0f, 0x22 / 255.0f, 0x3B / 255.0f, 1.0f),
                     "light viewer theme keeps selection text for checked toggle knob chrome");
    RequireColorNear(
        palette.selectionText, D2D1::ColorF(0x0E / 255.0f, 0x22 / 255.0f, 0x3B / 255.0f, 1.0f), "light viewer theme keeps the requested selection text chrome");
    Require(palette.scrollbarThumbHot.a > palette.scrollbarThumb.a, "light viewer theme keeps the hot scrollbar thumb stronger than the idle thumb");
    Require(palette.pressedFill.a > palette.hoverFill.a, "light viewer theme keeps pressed fill stronger than hover fill");
    RequireFloatNear(palette.selectionInactiveFill.a, 1.0f, 0.0001f, "light high-contrast viewer theme keeps inactive selection fully opaque");
}

void TestViewerThemePaletteDerivesDarkHighContrastChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF101214u;
    viewerTheme.textArgb                   = 0xFFE7E9EDu;
    viewerTheme.selectionBackgroundArgb    = 0xFF2A6DB2u;
    viewerTheme.selectionTextArgb          = 0xFFF8FAFCu;
    viewerTheme.accentArgb                 = 0xFF3A82D6u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = TRUE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette palette = MakeThemePaletteFromViewerTheme(viewerTheme);

    Require(palette.dark, "dark high-contrast viewer theme keeps the palette in dark mode");
    Require(palette.highContrast, "dark high-contrast viewer theme preserves high contrast");
    RequireColorNear(palette.hoverFill,
                     D2D1::ColorF(palette.focusStroke.r, palette.focusStroke.g, palette.focusStroke.b, 0.30f),
                     "dark high-contrast viewer theme strengthens hover fill alpha to the shared high-contrast contract");
    RequireColorNear(palette.pressedFill,
                     D2D1::ColorF(palette.focusStroke.r, palette.focusStroke.g, palette.focusStroke.b, 0.38f),
                     "dark high-contrast viewer theme strengthens pressed fill alpha to the shared high-contrast contract");
    RequireFloatNear(palette.selectionInactiveFill.a, 1.0f, 0.0001f, "dark high-contrast viewer theme keeps inactive selection fully opaque");
    Require(palette.pressedFill.a > palette.hoverFill.a, "dark high-contrast viewer theme keeps pressed fill stronger than hover fill");
}

void TestListIconColorUsesSelectionFillChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme   = MakeDefaultThemePalette(true);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.82f, 0.85f, 0.91f, 1.0f);

    RequireColorNear(ResolveListIconColor(theme, rowText, false),
                     BlendForTest(theme.selectionFill, rowText, 0.25f),
                     "nonselected list icons derive chrome from the palette selection fill");
    RequireColorNear(ResolveListIconColor(theme, rowText, true), rowText, "selected list icons keep row text chrome");
}

void TestListIconColorUsesViewerDerivedSelectionFillChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme   = MakeThemePaletteFromViewerTheme(viewerTheme);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.74f, 0.78f, 0.86f, 1.0f);

    RequireColorNear(ResolveListIconColor(theme, rowText, false),
                     BlendForTest(theme.selectionFill, rowText, 0.25f),
                     "viewer-derived nonselected list icons derive chrome from the palette selection fill");
    RequireColorDifferent(ResolveListIconColor(theme, rowText, false),
                          BlendForTest(theme.accent, rowText, 0.25f),
                          "viewer-derived nonselected list icons no longer read raw accent directly");
    RequireColorNear(ResolveListIconColor(theme, rowText, true), rowText, "viewer-derived selected list icons keep row text chrome");
}

void TestGridBusyColorUsesSelectionFillChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme   = MakeDefaultThemePalette(true);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.82f, 0.85f, 0.91f, 1.0f);

    RequireColorNear(ResolveGridBusyColor(theme, rowText, false), theme.selectionFill, "nonselected grid busy chrome derives from the palette selection fill");
    RequireColorNear(ResolveGridBusyColor(theme, rowText, true), rowText, "selected grid busy chrome keeps row text");
}

void TestGridBusyColorUsesViewerDerivedSelectionFillChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme   = MakeThemePaletteFromViewerTheme(viewerTheme);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.74f, 0.78f, 0.86f, 1.0f);

    RequireColorNear(ResolveGridBusyColor(theme, rowText, false),
                     theme.selectionFill,
                     "viewer-derived nonselected grid busy chrome derives from the palette selection fill");
    RequireColorDifferent(
        ResolveGridBusyColor(theme, rowText, false), theme.accent, "viewer-derived nonselected grid busy chrome no longer reads raw accent directly");
    RequireColorNear(ResolveGridBusyColor(theme, rowText, true), rowText, "viewer-derived selected grid busy chrome keeps row text");
}

void TestGridProgressVisualStyleUsesSharedTrackAndFillChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme   = MakeDefaultThemePalette(true);
    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.11f, 0.15f, 0.21f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.82f, 0.85f, 0.91f, 1.0f);

    const GridProgressVisualStyle unselected = ResolveGridProgressVisualStyle(theme, rowFill, rowText, false);
    RequireColorNear(unselected.track,
                     BlendForTest(theme.inputBorder, rowFill, 0.72f),
                     "unselected grid progress track derives from the palette input border plus row fill");
    RequireColorNear(unselected.fill, theme.selectionFill, "unselected grid progress fill derives from the palette selection fill");

    const GridProgressVisualStyle selected = ResolveGridProgressVisualStyle(theme, rowFill, rowText, true);
    RequireColorNear(selected.track, rowFill, "selected grid progress track stays on the selected row fill");
    RequireColorNear(selected.fill, rowText, "selected grid progress fill keeps row text chrome");
}

void TestGridProgressVisualStyleUsesViewerDerivedTrackAndFillChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme   = MakeThemePaletteFromViewerTheme(viewerTheme);
    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.15f, 0.19f, 0.26f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.74f, 0.78f, 0.86f, 1.0f);

    const GridProgressVisualStyle unselected = ResolveGridProgressVisualStyle(theme, rowFill, rowText, false);
    RequireColorNear(unselected.track,
                     BlendForTest(theme.inputBorder, rowFill, 0.72f),
                     "viewer-derived unselected grid progress track derives from the palette input border plus row fill");
    RequireColorDifferent(
        unselected.track, BlendForTest(theme.accent, rowFill, 0.72f), "viewer-derived unselected grid progress track no longer reads raw accent directly");
    RequireColorNear(unselected.fill, theme.selectionFill, "viewer-derived unselected grid progress fill derives from the palette selection fill");
    RequireColorDifferent(unselected.fill, theme.accent, "viewer-derived unselected grid progress fill no longer reads raw accent directly");

    const GridProgressVisualStyle selected = ResolveGridProgressVisualStyle(theme, rowFill, rowText, true);
    RequireColorNear(selected.track, rowFill, "viewer-derived selected grid progress track stays on the selected row fill");
    RequireColorNear(selected.fill, rowText, "viewer-derived selected grid progress fill keeps row text chrome");
}

void TestGridProgressVisualStyleUsesHighContrastTrackAndFillChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.80f, 0.84f, 0.90f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.10f, 0.12f, 0.16f, 1.0f);

    const GridProgressVisualStyle unselected = ResolveGridProgressVisualStyle(theme, rowFill, rowText, false);
    RequireColorNear(unselected.track,
                     BlendForTest(theme.inputBorder, rowFill, 0.72f),
                     "high-contrast unselected grid progress track derives from the palette input border plus row fill");
    RequireColorNear(unselected.fill, theme.selectionFill, "high-contrast unselected grid progress fill keeps palette selection chrome");

    const GridProgressVisualStyle selected = ResolveGridProgressVisualStyle(theme, rowFill, rowText, true);
    RequireColorNear(selected.track, rowFill, "high-contrast selected grid progress track stays on the selected row fill");
    RequireColorNear(selected.fill, rowText, "high-contrast selected grid progress fill keeps row text chrome");
}

void TestGridCheckboxVisualStyleUsesSharedCheckboxChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme   = MakeDefaultThemePalette(true);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.82f, 0.85f, 0.91f, 1.0f);

    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.13f, 0.42f, 0.73f, 1.0f);

    const GridCheckboxVisualStyle unselected = ResolveGridCheckboxVisualStyle(theme, rowFill, rowText, true, true, false, true);
    const CheckboxVisualStyle baseUnselected = ResolveCheckboxVisualStyle(theme, true, true, false, false, false, true);
    RequireColorNear(unselected.indicatorFill, baseUnselected.indicatorFill, "unselected grid checkbox fill matches the shared checkbox style");
    RequireColorNear(unselected.indicatorBorder, baseUnselected.indicatorBorder, "unselected grid checkbox border matches the shared checkbox style");
    RequireColorNear(unselected.check, baseUnselected.check, "unselected grid checkbox glyph matches the shared checkbox style");

    const GridCheckboxVisualStyle selected = ResolveGridCheckboxVisualStyle(theme, rowFill, rowText, true, true, true, true);
    RequireColorNear(selected.indicatorFill, rowText, "selected grid checkbox fill inverts to selected row text chrome");
    RequireColorNear(selected.indicatorBorder, rowText, "selected grid checkbox border inverts to selected row text chrome");
    RequireColorNear(selected.check, rowFill, "selected grid checkbox glyph inverts to selected row fill chrome");
}

void TestGridCheckboxVisualStyleUsesViewerDerivedSelectionTextChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme   = MakeThemePaletteFromViewerTheme(viewerTheme);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.74f, 0.78f, 0.86f, 1.0f);

    const D2D1_COLOR_F rowFill = theme.selectionFill;

    const GridCheckboxVisualStyle unselected = ResolveGridCheckboxVisualStyle(theme, rowFill, rowText, true, false, false, true);
    const CheckboxVisualStyle baseUnselected = ResolveCheckboxVisualStyle(theme, true, false, false, false, false, true);
    RequireColorNear(unselected.check, baseUnselected.check, "viewer-derived unselected grid checkbox glyph matches the shared checkbox style");

    const GridCheckboxVisualStyle selected = ResolveGridCheckboxVisualStyle(theme, rowFill, rowText, true, false, true, true);
    RequireColorNear(selected.indicatorFill, rowText, "viewer-derived selected grid checkbox fill inverts to selected row text chrome");
    RequireColorNear(selected.indicatorBorder, rowText, "viewer-derived selected grid checkbox border inverts to selected row text chrome");
    RequireColorNear(selected.check, rowFill, "viewer-derived selected grid checkbox glyph inverts to selected row fill chrome");
    RequireColorDifferent(selected.indicatorFill, theme.accent, "viewer-derived selected grid checkbox fill does not fall back to raw accent");
}

void TestGridSwatchVisualStyleUsesSharedRowChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme   = MakeDefaultThemePalette(true);
    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.18f, 0.20f, 0.24f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.82f, 0.85f, 0.91f, 1.0f);
    GridCellData swatchCell{};
    swatchCell.kind           = GridCellKind::ColorSwatch;
    swatchCell.hasSwatchValue = true;
    swatchCell.swatchArgb     = 0x8044AA33u;

    const GridSwatchVisualStyle selected   = ResolveGridSwatchVisualStyle(theme, rowFill, rowText, true, swatchCell);
    const GridSwatchVisualStyle unselected = ResolveGridSwatchVisualStyle(theme, rowFill, rowText, false, swatchCell);

    RequireColorNear(selected.fill,
                     BlendForTest(rowFill, D2D1::ColorF(0x44 / 255.0f, 0xAA / 255.0f, 0x33 / 255.0f, 1.0f), 0x80 / 255.0f),
                     "grid swatch fill resolves alpha against the row fill");
    RequireColorNear(selected.border, rowText, "selected grid swatch border inverts to selected row text chrome");
    RequireColorNear(
        unselected.border, BlendForTest(theme.inputBorder, rowFill, 0.12f), "unselected grid swatch border derives from the shared input border chrome");
}

void TestGridSwatchVisualStyleUsesViewerDerivedRowChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme   = MakeThemePaletteFromViewerTheme(viewerTheme);
    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.16f, 0.19f, 0.23f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.74f, 0.78f, 0.86f, 1.0f);
    GridCellData swatchCell{};
    swatchCell.kind           = GridCellKind::ColorSwatch;
    swatchCell.hasSwatchValue = true;
    swatchCell.swatchArgb     = 0xCC55B6E8u;

    const GridSwatchVisualStyle selected   = ResolveGridSwatchVisualStyle(theme, rowFill, rowText, true, swatchCell);
    const GridSwatchVisualStyle unselected = ResolveGridSwatchVisualStyle(theme, rowFill, rowText, false, swatchCell);

    RequireColorNear(selected.fill,
                     BlendForTest(rowFill, D2D1::ColorF(0x55 / 255.0f, 0xB6 / 255.0f, 0xE8 / 255.0f, 1.0f), 0xCC / 255.0f),
                     "viewer-derived grid swatch fill resolves alpha against the row fill");
    RequireColorNear(selected.border, rowText, "viewer-derived selected grid swatch border inverts to selected row text chrome");
    RequireColorNear(unselected.border,
                     BlendForTest(theme.inputBorder, rowFill, 0.12f),
                     "viewer-derived unselected grid swatch border derives from the shared input border chrome");
    RequireColorDifferent(
        unselected.border, BlendForTest(theme.accent, rowFill, 0.12f), "viewer-derived grid swatch border no longer reads raw accent directly");
}

void TestGridSwatchVisualStyleUsesHighContrastRowChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.82f, 0.86f, 0.92f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.09f, 0.11f, 0.15f, 1.0f);
    GridCellData swatchCell{};
    swatchCell.kind           = GridCellKind::ColorSwatch;
    swatchCell.hasSwatchValue = true;
    swatchCell.swatchArgb     = 0xCC3A79C9u;

    const GridSwatchVisualStyle selected   = ResolveGridSwatchVisualStyle(theme, rowFill, rowText, true, swatchCell);
    const GridSwatchVisualStyle unselected = ResolveGridSwatchVisualStyle(theme, rowFill, rowText, false, swatchCell);

    RequireColorNear(selected.fill,
                     BlendForTest(rowFill, D2D1::ColorF(0x3A / 255.0f, 0x79 / 255.0f, 0xC9 / 255.0f, 1.0f), 0xCC / 255.0f),
                     "high-contrast grid swatch fill resolves alpha against the row fill");
    RequireColorNear(selected.border, rowText, "high-contrast selected grid swatch border inverts to selected row text chrome");
    RequireColorNear(unselected.border,
                     BlendForTest(theme.inputBorder, rowFill, 0.12f),
                     "high-contrast unselected grid swatch border derives from shared input-border chrome");
}

void TestGridBadgeVisualStyleUsesSharedRowChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme   = MakeDefaultThemePalette(true);
    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.18f, 0.20f, 0.24f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.82f, 0.85f, 0.91f, 1.0f);

    const GridBadgeVisualStyle unselected = ResolveGridBadgeVisualStyle(theme, rowFill, rowText, false, AdornmentTone::Warning);
    const GridBadgeVisualStyle selected   = ResolveGridBadgeVisualStyle(theme, rowFill, rowText, true, AdornmentTone::Warning);

    RequireColorNear(unselected.fill, theme.warningFill, "grid badge uses the shared warning fill when the row is not selected");
    RequireColorNear(unselected.text, theme.warningText, "grid badge uses the shared warning text when the row is not selected");
    RequireColorNear(selected.fill, rowText, "selected grid badge fill inverts to selected row text chrome");
    RequireColorNear(selected.text, rowFill, "selected grid badge text inverts to selected row fill chrome");
}

void TestGridBadgeVisualStyleUsesViewerDerivedRowChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme   = MakeThemePaletteFromViewerTheme(viewerTheme);
    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.16f, 0.19f, 0.23f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.74f, 0.78f, 0.86f, 1.0f);

    const GridBadgeVisualStyle unselected = ResolveGridBadgeVisualStyle(theme, rowFill, rowText, false, AdornmentTone::Warning);
    const GridBadgeVisualStyle selected   = ResolveGridBadgeVisualStyle(theme, rowFill, rowText, true, AdornmentTone::Warning);

    RequireColorNear(unselected.fill, theme.warningFill, "viewer-derived grid badge uses the shared warning fill when the row is not selected");
    RequireColorNear(unselected.text, theme.warningText, "viewer-derived grid badge uses the shared warning text when the row is not selected");
    RequireColorNear(selected.fill, rowText, "viewer-derived selected grid badge fill inverts to selected row text chrome");
    RequireColorNear(selected.text, rowFill, "viewer-derived selected grid badge text inverts to selected row fill chrome");
    RequireColorDifferent(selected.fill, theme.accent, "viewer-derived selected grid badge fill no longer follows raw accent");
}

void TestGridBadgeVisualStyleUsesHighContrastRowChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.82f, 0.86f, 0.92f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.09f, 0.11f, 0.15f, 1.0f);

    const GridBadgeVisualStyle unselected = ResolveGridBadgeVisualStyle(theme, rowFill, rowText, false, AdornmentTone::Warning);
    const GridBadgeVisualStyle selected   = ResolveGridBadgeVisualStyle(theme, rowFill, rowText, true, AdornmentTone::Warning);

    RequireColorNear(unselected.fill, theme.warningFill, "high-contrast grid badge keeps shared warning fill when the row is not selected");
    RequireColorNear(unselected.text, theme.warningText, "high-contrast grid badge keeps shared warning text when the row is not selected");
    RequireColorNear(selected.fill, rowText, "high-contrast selected grid badge fill inverts to selected row text chrome");
    RequireColorNear(selected.text, rowFill, "high-contrast selected grid badge text inverts to selected row fill chrome");
}

void TestTreeBadgeVisualStyleUsesSharedAdornmentChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme           = MakeDefaultThemePalette(true);
    const TreeBadgeVisualStyle warning = ResolveTreeBadgeVisualStyle(theme, AdornmentTone::Warning);
    const TreeBadgeVisualStyle accent  = ResolveTreeBadgeVisualStyle(theme, AdornmentTone::Accent);

    RequireColorNear(warning.fill, theme.warningFill, "tree badge uses the shared warning fill");
    RequireColorNear(warning.text, theme.warningText, "tree badge uses the shared warning text");
    RequireColorNear(accent.fill, theme.selectionFill, "tree accent badge uses the shared selection fill");
    RequireColorNear(accent.text, theme.selectionText, "tree accent badge uses the shared selection text");
}

void TestTreeBadgeVisualStyleUsesViewerDerivedAdornmentChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme           = MakeThemePaletteFromViewerTheme(viewerTheme);
    const TreeBadgeVisualStyle warning = ResolveTreeBadgeVisualStyle(theme, AdornmentTone::Warning);
    const TreeBadgeVisualStyle accent  = ResolveTreeBadgeVisualStyle(theme, AdornmentTone::Accent);

    RequireColorNear(warning.fill, theme.warningFill, "viewer-derived tree badge uses the shared warning fill");
    RequireColorNear(warning.text, theme.warningText, "viewer-derived tree badge uses the shared warning text");
    RequireColorNear(accent.fill, theme.selectionFill, "viewer-derived tree accent badge uses the shared selection fill");
    RequireColorNear(accent.text, theme.selectionText, "viewer-derived tree accent badge uses the shared selection text");
    RequireColorDifferent(accent.fill, theme.accent, "viewer-derived tree accent badge no longer reads raw accent directly");
}

void TestTreeBadgeVisualStyleUsesHighContrastAdornmentChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const TreeBadgeVisualStyle warning = ResolveTreeBadgeVisualStyle(theme, AdornmentTone::Warning);
    const TreeBadgeVisualStyle accent  = ResolveTreeBadgeVisualStyle(theme, AdornmentTone::Accent);

    RequireColorNear(warning.fill, theme.warningFill, "high-contrast tree badge keeps shared warning fill");
    RequireColorNear(warning.text, theme.warningText, "high-contrast tree badge keeps shared warning text");
    RequireColorNear(accent.fill, theme.selectionFill, "high-contrast tree accent badge keeps shared selection fill");
    RequireColorNear(accent.text, theme.selectionText, "high-contrast tree accent badge keeps shared selection text");
}

void TestButtonVisualStyleMatchesPreferencesFlatChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme = MakeDefaultThemePalette(true);

    const ButtonVisualStyle idle = ResolveButtonVisualStyle(theme, true, false, false, false, false, false);
    RequireColorNear(idle.fill, theme.buttonFill, "idle button keeps neutral fill");
    Require(! idle.showBorder, "idle button hides border");
    Require(! idle.showFocus, "idle button hides focus ring");
    RequireColorNear(idle.text, theme.text, "idle button uses normal text color");

    const ButtonVisualStyle hovered = ResolveButtonVisualStyle(theme, true, true, false, false, false, false);
    Require(hovered.showBorder, "hovered button shows subtle border");
    RequireColorNear(hovered.fill, theme.buttonHotFill, "hovered button uses hot fill");
    RequireColorNear(hovered.border,
                     BlendForTest(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.22f : 0.16f),
                     "hovered button derives border chrome from the palette button border");

    const ButtonVisualStyle pressed = ResolveButtonVisualStyle(theme, true, false, true, false, false, false);
    Require(pressed.showBorder, "pressed button keeps border");
    RequireColorNear(pressed.fill, theme.buttonPressedFill, "pressed button uses pressed fill");
    RequireColorNear(pressed.border,
                     BlendForTest(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.24f : 0.18f),
                     "pressed button derives border chrome from the palette focus stroke");
    Require(pressed.textOffsetXDip > 0.0f && pressed.textOffsetYDip > 0.0f, "pressed button offsets text");

    const ButtonVisualStyle focused = ResolveButtonVisualStyle(theme, true, false, false, true, true, false);
    Require(focused.showBorder, "focused button shows border");
    Require(focused.showFocus, "focused button shows inner focus ring");
    RequireColorNear(focused.border,
                     BlendForTest(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.38f : 0.28f),
                     "focused button derives border chrome from the palette button border");

    const ButtonVisualStyle disabled = ResolveButtonVisualStyle(theme, false, false, false, false, false, false);
    Require(! disabled.showFocus, "disabled button has no focus ring");
    RequireColorNear(disabled.text, theme.disabledText, "disabled button uses disabled text color");
}

void TestPrimaryButtonVisualStyleUsesAccentChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme               = MakeDefaultThemePalette(true);
    const D2D1_COLOR_F expectedIdleFill    = BlendForTest(theme.buttonFill, theme.selectionFill, theme.dark ? (110.0f / 255.0f) : (90.0f / 255.0f));
    const D2D1_COLOR_F expectedHoverFill   = BlendForTest(expectedIdleFill, theme.selectionText, theme.dark ? (18.0f / 255.0f) : (12.0f / 255.0f));
    const D2D1_COLOR_F expectedPressedFill = BlendForTest(expectedIdleFill, theme.selectionText, theme.dark ? (24.0f / 255.0f) : (16.0f / 255.0f));

    const ButtonVisualStyle idle = ResolveButtonVisualStyle(theme, true, false, false, false, false, true);
    RequireColorNear(idle.fill, expectedIdleFill, "primary button uses muted selection-fill-tinted fill");
    RequireColorNear(idle.text, theme.selectionText, "primary button uses palette-derived selection text chrome");
    Require(! idle.showFocus, "idle primary button hides focus ring");
    Require(! idle.showBorder, "idle primary button hides border");

    const ButtonVisualStyle hovered = ResolveButtonVisualStyle(theme, true, true, false, false, false, true);
    RequireColorNear(hovered.fill, expectedHoverFill, "primary hover state darkens the muted selection fill with palette-derived selection text chrome");
    RequireColorNear(hovered.text, theme.selectionText, "primary hover state keeps palette-derived selection text chrome");
    Require(hovered.showBorder, "primary hover state shows border");
    RequireColorNear(hovered.border,
                     BlendForTest(theme.buttonBorder, theme.selectionText, theme.dark ? 0.24f : 0.18f),
                     "primary hover state derives border chrome from the palette selection text");

    const ButtonVisualStyle pressed = ResolveButtonVisualStyle(theme, true, false, true, false, false, true);
    RequireColorNear(pressed.fill, expectedPressedFill, "primary pressed state darkens the muted selection fill with palette-derived selection text chrome");
    RequireColorNear(pressed.text, theme.selectionText, "primary pressed state keeps palette-derived selection text chrome");
    RequireColorNear(pressed.border,
                     BlendForTest(theme.buttonBorder, theme.selectionText, theme.dark ? 0.34f : 0.26f),
                     "primary pressed state derives border chrome from the palette selection text");

    const ButtonVisualStyle focused = ResolveButtonVisualStyle(theme, true, false, false, true, true, true);
    RequireColorNear(focused.text, theme.selectionText, "primary focus state keeps palette-derived selection text chrome");
    Require(focused.showBorder, "keyboard-focused primary button shows border");
    Require(focused.showFocus, "keyboard-focused primary button shows focus ring");
    RequireColorNear(focused.border,
                     BlendForTest(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.40f : 0.30f),
                     "primary focus state derives border chrome from the palette button border");
    RequireColorNear(focused.focus,
                     BlendForTest(expectedIdleFill, theme.focusStroke, theme.dark ? 0.44f : 0.34f),
                     "primary focus state derives focus-ring chrome from the palette primary fill");

    const ButtonVisualStyle disabled = ResolveButtonVisualStyle(theme, false, false, false, false, false, true);
    RequireColorNear(disabled.text, theme.disabledText, "disabled primary button uses disabled text color");
}

void TestPrimaryButtonVisualStyleFallsBackToReadableText()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme  = MakeDefaultThemePalette(false);
    theme.buttonFill    = D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f);
    theme.selectionFill = D2D1::ColorF(0.58f, 0.78f, 0.92f, 1.0f);
    theme.selectionText = D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f);
    theme.text          = D2D1::ColorF(0.02f, 0.02f, 0.02f, 1.0f);

    const ButtonVisualStyle idle    = ResolveButtonVisualStyle(theme, true, false, false, false, false, true);
    const ButtonVisualStyle hovered = ResolveButtonVisualStyle(theme, true, true, false, false, false, true);
    const ButtonVisualStyle pressed = ResolveButtonVisualStyle(theme, true, false, true, false, false, true);
    const ButtonVisualStyle focused = ResolveButtonVisualStyle(theme, true, false, false, true, true, true);

    RequireColorNear(idle.text, theme.text, "low-contrast primary idle state falls back to readable palette text");
    RequireColorNear(hovered.text, theme.text, "low-contrast primary hover state falls back to readable palette text");
    RequireColorNear(pressed.text, theme.text, "low-contrast primary pressed state falls back to readable palette text");
    RequireColorNear(focused.text, theme.text, "low-contrast primary focus state falls back to readable palette text");
}

void TestButtonHighContrastFocusRingStaysVisibleWithoutKeyboardFocus()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const ButtonVisualStyle standardPointerFocused = ResolveButtonVisualStyle(theme, true, false, false, true, false, false);
    const ButtonVisualStyle primaryPointerFocused  = ResolveButtonVisualStyle(theme, true, false, false, true, false, true);
    const D2D1_COLOR_F expectedPrimaryIdleFill     = BlendForTest(theme.buttonFill, theme.selectionFill, theme.dark ? (110.0f / 255.0f) : (90.0f / 255.0f));

    Require(standardPointerFocused.showBorder, "high-contrast standard button keeps border visible without keyboard-focus gating");
    Require(standardPointerFocused.showFocus, "high-contrast standard button keeps focus ring visible without keyboard-focus gating");
    RequireColorNear(standardPointerFocused.border, theme.border, "high-contrast standard button focused border falls back to the palette border");
    RequireColorNear(standardPointerFocused.focus,
                     BlendForTest(theme.buttonFill, theme.focusStroke, theme.dark ? 0.52f : 0.38f),
                     "high-contrast standard button focus ring keeps palette focus chrome");

    Require(primaryPointerFocused.showBorder, "high-contrast primary button keeps border visible without keyboard-focus gating");
    Require(primaryPointerFocused.showFocus, "high-contrast primary button keeps focus ring visible without keyboard-focus gating");
    RequireColorNear(primaryPointerFocused.border, theme.border, "high-contrast primary button focused border falls back to the palette border");
    RequireColorNear(primaryPointerFocused.focus,
                     BlendForTest(expectedPrimaryIdleFill, theme.focusStroke, theme.dark ? 0.44f : 0.34f),
                     "high-contrast primary button focus ring keeps palette primary-fill chrome");
}

void TestButtonHighContrastDisabledBorderStaysVisible()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const ButtonVisualStyle standardDisabled = ResolveButtonVisualStyle(theme, false, false, false, false, false, false);
    const ButtonVisualStyle primaryDisabled  = ResolveButtonVisualStyle(theme, false, false, false, false, false, true);

    Require(standardDisabled.showBorder, "high-contrast disabled standard button keeps border visible");
    Require(! standardDisabled.showFocus, "high-contrast disabled standard button still hides focus ring");
    RequireColorNear(standardDisabled.border, theme.border, "high-contrast disabled standard button uses the palette border");
    RequireColorNear(standardDisabled.text, theme.disabledText, "high-contrast disabled standard button keeps disabled text color");

    Require(primaryDisabled.showBorder, "high-contrast disabled primary button keeps border visible");
    Require(! primaryDisabled.showFocus, "high-contrast disabled primary button still hides focus ring");
    RequireColorNear(primaryDisabled.border, theme.border, "high-contrast disabled primary button uses the palette border");
    RequireColorNear(primaryDisabled.text, theme.disabledText, "high-contrast disabled primary button keeps disabled text color");
}

void TestButtonVisualStyleInterpolatesHoverAndFocusStrength()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme          = MakeDefaultThemePalette(true);
    const ButtonVisualStyle halfHover = ResolveButtonVisualStyle(theme, true, true, false, false, false, false, 0.5f, 0.0f);
    RequireColorNear(halfHover.fill, BlendForTest(theme.buttonFill, theme.buttonHotFill, 0.5f), "button hover animation blends halfway toward hot fill");
    Require(halfHover.showBorder, "button hover animation enables border chrome while in flight");

    const ButtonVisualStyle focused   = ResolveButtonVisualStyle(theme, true, false, false, true, true, false);
    const ButtonVisualStyle halfFocus = ResolveButtonVisualStyle(theme, true, false, false, true, true, false, 0.0f, 0.5f);
    Require(halfFocus.showFocus, "button focus animation keeps focus chrome visible while in flight");
    RequireFloatNear(halfFocus.focus.a, focused.focus.a * 0.5f, 0.0001f, "button focus animation scales focus-ring alpha by progress");
}

void TestPrimaryButtonVisualStyleUsesViewerDerivedSelectionTextChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF2D6FB7u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme            = MakeThemePaletteFromViewerTheme(viewerTheme);
    const ButtonVisualStyle idle        = ResolveButtonVisualStyle(theme, true, false, false, false, false, true);
    const ButtonVisualStyle hovered     = ResolveButtonVisualStyle(theme, true, true, false, false, false, true);
    const ButtonVisualStyle pressed     = ResolveButtonVisualStyle(theme, true, false, true, false, false, true);
    const D2D1_COLOR_F expectedIdleFill = BlendForTest(theme.buttonFill, theme.selectionFill, theme.dark ? (110.0f / 255.0f) : (90.0f / 255.0f));

    RequireColorNear(idle.fill, expectedIdleFill, "viewer-derived primary button idle fill uses the shared selection fill");
    RequireColorNear(idle.text, theme.selectionText, "viewer-derived primary button idle chrome uses the palette selection text");
    RequireColorNear(hovered.fill,
                     BlendForTest(expectedIdleFill, theme.selectionText, theme.dark ? (18.0f / 255.0f) : (12.0f / 255.0f)),
                     "viewer-derived primary button hover fill uses palette selection text chrome");
    RequireColorNear(hovered.border,
                     BlendForTest(theme.buttonBorder, theme.selectionText, theme.dark ? 0.24f : 0.18f),
                     "viewer-derived primary button hover border uses the palette selection text");
    RequireColorNear(pressed.fill,
                     BlendForTest(expectedIdleFill, theme.selectionText, theme.dark ? (24.0f / 255.0f) : (16.0f / 255.0f)),
                     "viewer-derived primary button pressed fill uses palette selection text chrome");
    RequireColorNear(pressed.border,
                     BlendForTest(theme.buttonBorder, theme.selectionText, theme.dark ? 0.34f : 0.26f),
                     "viewer-derived primary button pressed border uses the palette selection text");
    RequireColorNear(pressed.text, theme.selectionText, "viewer-derived primary button pressed chrome uses the palette selection text");
    const ButtonVisualStyle focused = ResolveButtonVisualStyle(theme, true, false, false, true, true, true);
    RequireColorNear(focused.focus,
                     BlendForTest(expectedIdleFill, theme.focusStroke, theme.dark ? 0.44f : 0.34f),
                     "viewer-derived primary button focus ring uses the palette primary fill");
    RequireColorDifferent(idle.fill, theme.accent, "viewer-derived primary button idle fill follows selection fill instead of raw accent");
    RequireColorDifferent(idle.text, theme.text, "viewer-derived primary button keeps accent-surface text distinct from neutral text when requested");
}

void TestButtonVisualStyleUsesViewerDerivedPressedBorderChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF4F8EDCu;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme        = MakeThemePaletteFromViewerTheme(viewerTheme);
    const ButtonVisualStyle pressed = ResolveButtonVisualStyle(theme, true, false, true, false, false, false);

    RequireColorNear(pressed.border,
                     BlendForTest(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.24f : 0.18f),
                     "viewer-derived pressed button border chrome uses the palette focus stroke");
}

void TestButtonVisualStyleUsesViewerDerivedHotAndPressedFillChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF4F8EDCu;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme        = MakeThemePaletteFromViewerTheme(viewerTheme);
    const ButtonVisualStyle hovered = ResolveButtonVisualStyle(theme, true, true, false, false, false, false);
    const ButtonVisualStyle pressed = ResolveButtonVisualStyle(theme, true, false, true, false, false, false);

    RequireColorNear(hovered.fill,
                     BlendForTest(theme.buttonFill, theme.focusStroke, theme.dark ? 0.10f : 0.05f),
                     "viewer-derived hovered button fill chrome uses the palette focus stroke");
    RequireColorNear(pressed.fill,
                     BlendForTest(theme.buttonFill, theme.focusStroke, theme.dark ? 0.18f : 0.10f),
                     "viewer-derived pressed button fill chrome uses the palette focus stroke");
}

void TestAdornmentColorsUsePaletteChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme = MakeDefaultThemePalette(true);

    D2D1_COLOR_F fill{};
    D2D1_COLOR_F text{};

    ResolveAdornmentColors(theme, AdornmentTone::Accent, fill, text);
    RequireColorNear(fill, theme.selectionFill, "accent adornment uses shared selection fill");
    RequireColorNear(text, theme.selectionText, "accent adornment uses palette-derived selection text");

    ResolveAdornmentColors(theme, AdornmentTone::Info, fill, text);
    RequireColorNear(fill, theme.infoFill, "info adornment uses info fill");
    RequireColorNear(text, theme.infoText, "info adornment uses info text");

    ResolveAdornmentColors(theme, AdornmentTone::Warning, fill, text);
    RequireColorNear(fill, theme.warningFill, "warning adornment uses warning fill");
    RequireColorNear(text, theme.warningText, "warning adornment uses warning text");

    ResolveAdornmentColors(theme, AdornmentTone::Error, fill, text);
    RequireColorNear(fill, theme.errorFill, "error adornment uses error fill");
    RequireColorNear(text, theme.errorText, "error adornment uses error text");
}

void TestAdornmentColorsUseViewerDerivedSelectionAndAlertChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF2D6FB7u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme = MakeThemePaletteFromViewerTheme(viewerTheme);

    D2D1_COLOR_F fill{};
    D2D1_COLOR_F text{};

    ResolveAdornmentColors(theme, AdornmentTone::Accent, fill, text);
    RequireColorNear(fill, theme.selectionFill, "viewer-derived accent adornment uses shared selection fill");
    RequireColorNear(text, theme.selectionText, "viewer-derived accent adornment uses palette selection text");
    RequireColorDifferent(fill, theme.accent, "viewer-derived accent adornment follows selection fill instead of raw accent");

    ResolveAdornmentColors(theme, AdornmentTone::Info, fill, text);
    RequireColorNear(fill, theme.infoFill, "viewer-derived info adornment uses info fill");
    RequireColorNear(text, theme.infoText, "viewer-derived info adornment uses info text");

    ResolveAdornmentColors(theme, AdornmentTone::Warning, fill, text);
    RequireColorNear(fill, theme.warningFill, "viewer-derived warning adornment uses warning fill");
    RequireColorNear(text, theme.warningText, "viewer-derived warning adornment uses warning text");

    ResolveAdornmentColors(theme, AdornmentTone::Error, fill, text);
    RequireColorNear(fill, theme.errorFill, "viewer-derived error adornment uses error fill");
    RequireColorNear(text, theme.errorText, "viewer-derived error adornment uses error text");
}

void TestToggleVisualStyleMatchesPreferencesSwitchChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme = MakeDefaultThemePalette(true);

    const ToggleVisualStyle idle = ResolveToggleVisualStyle(theme, true, false, false, false, false, false);
    Require(! idle.showRowFill, "idle switch keeps neutral row treatment");
    RequireColorNear(idle.trackFill, BlendForTest(theme.inputFill, theme.border, theme.dark ? 0.18f : 0.08f), "idle switch uses neutral track");
    RequireColorNear(idle.knobFill, theme.toggleKnobFill, "idle switch uses palette-derived neutral knob chrome");
    RequireColorNear(
        idle.knobBorder, BlendForTest(idle.knobFill, idle.trackBorder, 0.28f), "idle switch resolves knob border chrome from knob and track colors");

    const ToggleVisualStyle hovered = ResolveToggleVisualStyle(theme, true, true, false, false, false, false);
    Require(hovered.showRowFill, "hovered switch tints row softly");
    RequireColorNear(
        hovered.rowFill,
        BlendForTest(theme.surfaceBackground, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.09f : 0.05f),
        "hovered switch uses the shared hover overlay token for row tint");

    const ToggleVisualStyle pressed = ResolveToggleVisualStyle(theme, true, false, true, false, false, false);
    Require(pressed.showRowFill, "pressed switch keeps row tint visible");
    RequireColorNear(
        pressed.rowFill,
        BlendForTest(theme.surfaceBackground, D2D1::ColorF(theme.pressedFill.r, theme.pressedFill.g, theme.pressedFill.b, 1.0f), theme.dark ? 0.16f : 0.10f),
        "pressed switch uses the shared pressed overlay token for row tint");

    const ToggleVisualStyle checked = ResolveToggleVisualStyle(theme, true, false, false, false, false, true);
    RequireColorNear(checked.trackFill, theme.selectionFill, "checked switch uses shared selection fill for track chrome");
    RequireColorNear(checked.trackBorder,
                     BlendForTest(theme.selectionFill, theme.selectionText, theme.dark ? 0.18f : 0.10f),
                     "checked switch uses selection-fill and selection-text track border chrome");
    RequireColorNear(checked.knobFill, theme.toggleKnobCheckedFill, "checked switch uses palette-derived checked knob chrome");
    RequireColorNear(checked.knobBorder,
                     BlendForTest(checked.knobFill, checked.trackBorder, 0.28f),
                     "checked switch resolves knob border chrome from checked knob and track colors");

    const ToggleVisualStyle pointerFocused = ResolveToggleVisualStyle(theme, true, false, false, true, false, false);
    Require(! pointerFocused.showFocus, "pointer-focused switch keeps focus chrome quiet");

    const ToggleVisualStyle keyboardFocused = ResolveToggleVisualStyle(theme, true, false, false, true, true, false);
    Require(keyboardFocused.showFocus, "keyboard-focused switch shows focus chrome");

    ThemePalette highContrastTheme              = theme;
    highContrastTheme.highContrast              = true;
    const ToggleVisualStyle highContrastFocused = ResolveToggleVisualStyle(highContrastTheme, true, false, false, true, false, false);
    Require(highContrastFocused.showFocus, "high-contrast switch keeps focus chrome visible even without keyboard focus visibility");

    const ToggleVisualStyle highContrastChecked = ResolveToggleVisualStyle(highContrastTheme, true, false, false, false, false, true);
    RequireColorNear(highContrastChecked.trackBorder,
                     BlendForTest(highContrastTheme.selectionFill, highContrastTheme.selectionText, theme.dark ? 0.18f : 0.10f),
                     "high-contrast checked switch keeps the shared checked track border chrome");
    RequireColorNear(highContrastChecked.knobBorder,
                     BlendForTest(highContrastChecked.knobFill, highContrastChecked.trackBorder, 0.28f),
                     "high-contrast checked switch keeps knob border chrome derived from the resolved checked track border");

    const ToggleVisualStyle disabledChecked = ResolveToggleVisualStyle(theme, false, false, false, false, false, true);
    RequireColorNear(disabledChecked.text, theme.disabledText, "disabled switch uses disabled text color");
}

void TestToggleVisualStyleUsesViewerDerivedKnobChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF2D6FB7u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme        = MakeThemePaletteFromViewerTheme(viewerTheme);
    const ToggleVisualStyle idle    = ResolveToggleVisualStyle(theme, true, false, false, false, false, false);
    const ToggleVisualStyle hovered = ResolveToggleVisualStyle(theme, true, true, false, false, false, false);
    const ToggleVisualStyle pressed = ResolveToggleVisualStyle(theme, true, false, true, false, false, false);
    const ToggleVisualStyle checked = ResolveToggleVisualStyle(theme, true, false, false, false, false, true);

    RequireColorNear(idle.trackBorder, theme.inputBorder, "viewer-derived toggle idle border chrome uses the palette input border");
    RequireColorNear(idle.knobFill, theme.toggleKnobFill, "viewer-derived toggle idle chrome uses the palette neutral knob");
    RequireColorNear(idle.knobBorder,
                     BlendForTest(idle.knobFill, idle.trackBorder, 0.28f),
                     "viewer-derived toggle idle knob border chrome uses the resolved knob and track colors");
    RequireColorNear(
        hovered.rowFill,
        BlendForTest(theme.surfaceBackground, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.09f : 0.05f),
        "viewer-derived toggle hover row tint uses the shared hover overlay token");
    RequireColorNear(
        pressed.rowFill,
        BlendForTest(theme.surfaceBackground, D2D1::ColorF(theme.pressedFill.r, theme.pressedFill.g, theme.pressedFill.b, 1.0f), theme.dark ? 0.16f : 0.10f),
        "viewer-derived toggle pressed row tint uses the shared pressed overlay token");
    RequireColorNear(checked.trackFill, theme.selectionFill, "viewer-derived toggle checked fill uses the shared selection fill");
    RequireColorNear(checked.trackBorder,
                     BlendForTest(theme.selectionFill, theme.selectionText, theme.dark ? 0.18f : 0.10f),
                     "viewer-derived toggle checked border chrome uses shared selection fill and text");
    RequireColorNear(checked.knobFill, theme.toggleKnobCheckedFill, "viewer-derived toggle checked chrome uses the palette checked knob");
    RequireColorNear(checked.knobBorder,
                     BlendForTest(checked.knobFill, checked.trackBorder, 0.28f),
                     "viewer-derived toggle checked knob border chrome uses the resolved knob and track colors");
    RequireColorDifferent(checked.trackFill, theme.accent, "viewer-derived toggle checked fill follows selection fill instead of raw accent");
    RequireColorDifferent(idle.knobFill, checked.knobFill, "viewer-derived toggle keeps neutral and checked knob chrome distinct");
}

void TestToggleHighContrastDisabledBordersStayVisible()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(true);
    theme.highContrast = true;

    const ToggleVisualStyle disabledUnchecked = ResolveToggleVisualStyle(theme, false, false, false, false, false, false);
    const ToggleVisualStyle disabledChecked   = ResolveToggleVisualStyle(theme, false, false, false, false, false, true);

    RequireColorNear(disabledUnchecked.trackBorder,
                     BlendForTest(theme.windowBackground, theme.inputBorder, theme.dark ? 0.58f : 0.42f),
                     "high-contrast disabled unchecked toggle keeps the shared disabled input-border chrome");
    RequireColorNear(disabledUnchecked.knobBorder,
                     BlendForTest(disabledUnchecked.knobFill, disabledUnchecked.trackBorder, 0.28f),
                     "high-contrast disabled unchecked toggle keeps knob-border chrome derived from the disabled track border");
    RequireColorNear(disabledUnchecked.text, theme.disabledText, "high-contrast disabled unchecked toggle keeps disabled text color");

    RequireColorNear(disabledChecked.trackBorder,
                     BlendForTest(theme.selectionFill, theme.selectionText, theme.dark ? 0.18f : 0.10f),
                     "high-contrast disabled checked toggle keeps the resolved checked-track border chrome");
    RequireColorNear(disabledChecked.knobBorder,
                     BlendForTest(disabledChecked.knobFill, disabledChecked.trackBorder, 0.28f),
                     "high-contrast disabled checked toggle keeps knob-border chrome derived from the checked track border");
    RequireColorNear(disabledChecked.text, theme.disabledText, "high-contrast disabled checked toggle keeps disabled text color");
    RequireColorDifferent(
        disabledChecked.trackBorder, disabledUnchecked.trackBorder, "high-contrast disabled checked toggle keeps distinct checked-track border chrome");
}

void TestToggleVisualStyleInterpolatesHoverStrength()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme          = MakeDefaultThemePalette(true);
    const ToggleVisualStyle halfHover = ResolveToggleVisualStyle(theme, true, true, false, false, false, false, 0.5f, 0.0f);
    const D2D1_COLOR_F expectedRowFill =
        BlendForTest(theme.surfaceBackground, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), (theme.dark ? 0.09f : 0.05f) * 0.5f);
    Require(halfHover.showRowFill, "toggle hover animation keeps row tint visible while in flight");
    RequireColorNear(halfHover.rowFill, expectedRowFill, "toggle hover animation blends row tint by progress");
}

void TestToggleLayoutMetricsUseCompactSwitchOnlyChromeWhenTextIsEmpty()
{
    using namespace RedSalamander::DxUi;

    Toggle toggle;
    toggle.SetBounds(D2D1::RectF(0.0f, 0.0f, 60.0f, 32.0f));

    const ToggleLayoutMetrics metrics = toggle.GetLayoutMetrics();
    Require(metrics.compactSwitchOnly, "empty-text toggle uses compact switch-only layout");
    RequireRectHasArea(metrics.trackRect, "empty-text toggle exposes a track rect");
    Require(metrics.textRect.right <= metrics.textRect.left + 0.5f, "empty-text toggle collapses the text rect");
    Require((metrics.backgroundRect.right - metrics.backgroundRect.left) < 57.0f, "empty-text toggle keeps hover chrome narrower than the full host");

    const float trackCenterX = (metrics.trackRect.left + metrics.trackRect.right) * 0.5f;
    RequireFloatNear(trackCenterX, 30.0f, 1.0f, "empty-text toggle centers the switch track within the host");
}

void TestColorSwatchVisualStyleUsesSharedOverlayChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme                      = MakeDefaultThemePalette(true);
    const ColorSwatchVisualStyle idle             = ResolveColorSwatchVisualStyle(theme, true, false, false, false);
    const ColorSwatchVisualStyle hovered          = ResolveColorSwatchVisualStyle(theme, true, true, false, false);
    const ColorSwatchVisualStyle pressed          = ResolveColorSwatchVisualStyle(theme, true, true, true, false);
    const ColorSwatchVisualStyle focusedHovered   = ResolveColorSwatchVisualStyle(theme, true, true, false, true);
    const ColorSwatchVisualStyle disabled         = ResolveColorSwatchVisualStyle(theme, false, false, false, false);
    const ColorSwatchVisualStyle assigned         = ResolveColorSwatchVisualStyle(theme, true, false, false, false, false, 0x8044AA33u);
    const ColorSwatchVisualStyle disabledAssigned = ResolveColorSwatchVisualStyle(theme, false, false, false, false, false, 0x8044AA33u);
    const D2D1_COLOR_F assignedBase               = BlendForTest(idle.fill, D2D1::ColorF(0x44 / 255.0f, 0xAA / 255.0f, 0x33 / 255.0f, 1.0f), 0x80 / 255.0f);
    const D2D1_COLOR_F disabledAssignedBase       = BlendForTest(disabled.fill, D2D1::ColorF(0x44 / 255.0f, 0xAA / 255.0f, 0x33 / 255.0f, 1.0f), 0x80 / 255.0f);

    RequireColorNear(idle.fill, theme.inputFill, "idle color swatch uses the palette input fill");
    RequireColorNear(idle.swatchFill, idle.fill, "empty color swatch inner fill falls back to the shared container fill");
    RequireColorNear(idle.swatchBorder, BlendForTest(idle.border, idle.fill, 0.18f), "empty color swatch inner border derives from the resolved outer chrome");
    RequireColorNear(hovered.fill,
                     BlendForTest(idle.fill, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.08f : 0.04f),
                     "hovered color swatch uses the shared hover overlay token");
    RequireColorNear(pressed.fill,
                     BlendForTest(idle.fill, D2D1::ColorF(theme.pressedFill.r, theme.pressedFill.g, theme.pressedFill.b, 1.0f), theme.dark ? 0.14f : 0.08f),
                     "pressed color swatch uses the shared pressed overlay token");
    RequireColorNear(hovered.border,
                     BlendForTest(theme.inputBorder, theme.focusStroke, theme.dark ? 0.28f : 0.20f),
                     "hovered color swatch border derives chrome from the palette input border");
    RequireColorNear(focusedHovered.border, theme.focusStroke, "focused color swatch keeps the palette focus stroke");
    RequireColorNear(focusedHovered.focus, theme.focusStroke, "focused color swatch uses the palette focus stroke for the focus ring");
    RequireColorNear(disabled.fill,
                     BlendForTest(theme.windowBackground, theme.inputFill, theme.dark ? 0.78f : 0.62f),
                     "disabled color swatch mutes fill toward the neutral palette");
    RequireColorNear(assigned.swatchFill, assignedBase, "assigned color swatch inner fill resolves alpha against the shared container fill");
    RequireColorNear(assigned.swatchBorder,
                     BlendForTest(assigned.border, assigned.swatchFill, 0.28f),
                     "assigned color swatch inner border derives from the resolved outer border");
    RequireColorNear(disabledAssigned.swatchFill,
                     BlendForTest(disabled.fill, disabledAssignedBase, theme.dark ? 0.52f : 0.44f),
                     "disabled assigned color swatch mutes inner fill through the resolved style contract");
}

void TestColorSwatchUsesViewerDerivedOverlayChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme                      = MakeThemePaletteFromViewerTheme(viewerTheme);
    const ColorSwatchVisualStyle idle             = ResolveColorSwatchVisualStyle(theme, true, false, false, false);
    const ColorSwatchVisualStyle hovered          = ResolveColorSwatchVisualStyle(theme, true, true, false, false);
    const ColorSwatchVisualStyle pressed          = ResolveColorSwatchVisualStyle(theme, true, true, true, false);
    const ColorSwatchVisualStyle focused          = ResolveColorSwatchVisualStyle(theme, true, false, false, true);
    const ColorSwatchVisualStyle assigned         = ResolveColorSwatchVisualStyle(theme, true, false, false, false, false, 0xCC55B6E8u);
    const ColorSwatchVisualStyle disabledAssigned = ResolveColorSwatchVisualStyle(theme, false, false, false, false, false, 0xCC55B6E8u);
    const D2D1_COLOR_F assignedBase               = BlendForTest(idle.fill, D2D1::ColorF(0x55 / 255.0f, 0xB6 / 255.0f, 0xE8 / 255.0f, 1.0f), 0xCC / 255.0f);
    const D2D1_COLOR_F disabledAssignedBase =
        BlendForTest(disabledAssigned.fill, D2D1::ColorF(0x55 / 255.0f, 0xB6 / 255.0f, 0xE8 / 255.0f, 1.0f), 0xCC / 255.0f);

    RequireColorNear(idle.border, theme.inputBorder, "viewer-derived color swatch idle border uses the palette input border");
    RequireColorNear(hovered.fill,
                     BlendForTest(idle.fill, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.08f : 0.04f),
                     "viewer-derived color swatch hover fill uses the shared hover overlay token");
    RequireColorNear(pressed.fill,
                     BlendForTest(idle.fill, D2D1::ColorF(theme.pressedFill.r, theme.pressedFill.g, theme.pressedFill.b, 1.0f), theme.dark ? 0.14f : 0.08f),
                     "viewer-derived color swatch pressed fill uses the shared pressed overlay token");
    RequireColorNear(focused.border, theme.focusStroke, "viewer-derived color swatch focus border uses the palette focus stroke");
    RequireColorNear(focused.focus, theme.focusStroke, "viewer-derived color swatch focus ring uses the palette focus stroke");
    RequireColorNear(assigned.swatchFill, assignedBase, "viewer-derived color swatch inner fill resolves alpha against the shared container fill");
    RequireColorNear(assigned.swatchBorder,
                     BlendForTest(assigned.border, assigned.swatchFill, 0.28f),
                     "viewer-derived color swatch inner border derives from the resolved border chrome");
    RequireColorNear(disabledAssigned.swatchFill,
                     BlendForTest(disabledAssigned.fill, disabledAssignedBase, theme.dark ? 0.52f : 0.44f),
                     "viewer-derived disabled color swatch mutes inner fill through the shared style contract");
}

void TestColorSwatchHighContrastFocusRingStaysVisibleWithoutKeyboardFocus()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const ColorSwatchVisualStyle pointerFocused  = ResolveColorSwatchVisualStyle(theme, true, false, false, true, false);
    const ColorSwatchVisualStyle keyboardFocused = ResolveColorSwatchVisualStyle(theme, true, false, false, true, true);

    Require(pointerFocused.showFocus, "high-contrast color swatch keeps focus ring visible without keyboard-focus gating");
    Require(keyboardFocused.showFocus, "keyboard-focused color swatch keeps focus ring visible");
    RequireColorNear(pointerFocused.focus, theme.focusStroke, "high-contrast color swatch focus ring uses the palette focus stroke");
    RequireColorNear(pointerFocused.border, theme.focusStroke, "high-contrast color swatch focused border still uses the palette focus stroke");
}

void TestCardPanelVisualStyleUsesSharedSurfaceChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme         = MakeDefaultThemePalette(true);
    const CardPanelVisualStyle style = ResolveCardPanelVisualStyle(theme);

    RequireColorNear(style.fill, theme.cardBackground, "card panel uses the palette card background token");
    RequireColorNear(style.border, theme.borderDefault, "card panel border uses the palette border default token");
}

void TestCardPanelUsesViewerDerivedSurfaceChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFFF7F4EEu;
    viewerTheme.textArgb                   = 0xFF24292Fu;
    viewerTheme.selectionBackgroundArgb    = 0xFFD9E8FFu;
    viewerTheme.selectionTextArgb          = 0xFF0E223Bu;
    viewerTheme.accentArgb                 = 0xFFB85C1Eu;
    viewerTheme.alertErrorBackgroundArgb   = 0xFFFFE2E0u;
    viewerTheme.alertErrorTextArgb         = 0xFF7D2018u;
    viewerTheme.alertWarningBackgroundArgb = 0xFFFFF1D2u;
    viewerTheme.alertWarningTextArgb       = 0xFF815100u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFFE2F0FFu;
    viewerTheme.alertInfoTextArgb          = 0xFF194A78u;
    viewerTheme.darkMode                   = FALSE;
    viewerTheme.highContrast               = TRUE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = FALSE;

    const ThemePalette theme         = MakeThemePaletteFromViewerTheme(viewerTheme);
    const CardPanelVisualStyle style = ResolveCardPanelVisualStyle(theme);

    RequireColorNear(style.fill, theme.cardBackground, "viewer-derived card panel uses the palette card background token");
    RequireColorNear(style.border, theme.borderDefault, "viewer-derived card panel border uses the palette border default token");
    RequireColorDifferent(style.border, theme.accent, "viewer-derived card panel border follows shared surface chrome instead of raw accent");
}

void TestCardPanelHighContrastUsesPaletteSurfaceChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const CardPanelVisualStyle style = ResolveCardPanelVisualStyle(theme);

    RequireColorNear(style.fill, theme.cardBackground, "high-contrast card panel keeps the palette card background token");
    RequireColorNear(style.border, theme.borderDefault, "high-contrast card panel keeps the palette border default token");
}

void TestGridSurfaceVisualStyleUsesSharedGridChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme           = MakeDefaultThemePalette(true);
    const GridSurfaceVisualStyle style = ResolveGridSurfaceVisualStyle(theme);

    RequireColorNear(style.fill, theme.surfaceBackground, "grid surface style uses the palette surface background");
    RequireColorNear(style.border, theme.border, "grid surface style uses the palette border");
    RequireColorNear(style.headerFill, theme.headerBackground, "grid surface style uses the palette header background");
    RequireColorNear(style.headerBorder, theme.border, "grid surface style header separator uses the palette border");
    RequireColorNear(style.rowSeparator, theme.gridLine, "grid surface style row separators use the palette grid line");
    RequireColorNear(style.columnSeparator, theme.gridLine, "grid surface style column separators use the palette grid line");
    RequireColorNear(style.emptyText, theme.subduedText, "grid surface style empty text uses the palette subdued text");
}

void TestGridSurfaceVisualStyleUsesViewerDerivedGridChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF10151Bu;
    viewerTheme.textArgb                   = 0xFFE8EEF9u;
    viewerTheme.selectionBackgroundArgb    = 0xFF3F74C7u;
    viewerTheme.selectionTextArgb          = 0xFFF6FAFFu;
    viewerTheme.accentArgb                 = 0xFFD15A7Au;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5A1D24u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD9DDu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF5A430Eu;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE4A1u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF173149u;
    viewerTheme.alertInfoTextArgb          = 0xFFD7E9FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme           = MakeThemePaletteFromViewerTheme(viewerTheme);
    const GridSurfaceVisualStyle style = ResolveGridSurfaceVisualStyle(theme);

    RequireColorNear(style.fill, theme.surfaceBackground, "viewer-derived grid surface style uses the palette surface background");
    RequireColorNear(style.border, theme.border, "viewer-derived grid surface style uses the palette border");
    RequireColorNear(style.headerFill, theme.headerBackground, "viewer-derived grid surface style uses the palette header background");
    RequireColorNear(style.headerBorder, theme.border, "viewer-derived grid surface style header separator uses the palette border");
    RequireColorNear(style.rowSeparator, theme.gridLine, "viewer-derived grid surface style row separators use the palette grid line");
    RequireColorNear(style.columnSeparator, theme.gridLine, "viewer-derived grid surface style column separators use the palette grid line");
    RequireColorNear(style.emptyText, theme.subduedText, "viewer-derived grid surface style empty text uses the palette subdued text");
    RequireColorDifferent(style.fill, theme.accent, "viewer-derived grid surface style follows grid chrome instead of raw accent");
    RequireColorDifferent(style.rowSeparator, theme.accent, "viewer-derived grid surface style row separators do not read raw accent");
    RequireColorDifferent(style.columnSeparator, theme.accent, "viewer-derived grid surface style column separators do not read raw accent");
}

void TestGridSurfaceHighContrastUsesPaletteGridChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const GridSurfaceVisualStyle style = ResolveGridSurfaceVisualStyle(theme);

    RequireColorNear(style.fill, theme.surfaceBackground, "high-contrast grid surface style keeps the palette surface background");
    RequireColorNear(style.border, theme.border, "high-contrast grid surface style keeps the palette border");
    RequireColorNear(style.headerFill, theme.headerBackground, "high-contrast grid surface style keeps the palette header background");
    RequireColorNear(style.headerBorder, theme.border, "high-contrast grid surface style keeps the palette header separator border");
    RequireColorNear(style.rowSeparator, theme.gridLine, "high-contrast grid surface style keeps the palette grid line for rows");
    RequireColorNear(style.columnSeparator, theme.gridLine, "high-contrast grid surface style keeps the palette grid line for columns");
    RequireColorNear(style.emptyText, theme.subduedText, "high-contrast grid surface style keeps the palette subdued text");
}

void TestTreeSurfaceVisualStyleUsesSharedTreeChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme           = MakeDefaultThemePalette(true);
    const TreeSurfaceVisualStyle style = ResolveTreeSurfaceVisualStyle(theme);

    RequireColorNear(style.fill, theme.surfaceBackground, "tree surface style uses the palette surface background");
    RequireColorNear(style.border, theme.border, "tree surface style uses the palette border");
    RequireColorNear(style.emptyText, theme.subduedText, "tree surface style empty text uses the palette subdued text");
}

void TestTreeSurfaceVisualStyleUsesViewerDerivedTreeChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFFF4F0E7u;
    viewerTheme.textArgb                   = 0xFF252B31u;
    viewerTheme.selectionBackgroundArgb    = 0xFFDDE9FFu;
    viewerTheme.selectionTextArgb          = 0xFF10233Au;
    viewerTheme.accentArgb                 = 0xFFB86A29u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFFFFE1DDu;
    viewerTheme.alertErrorTextArgb         = 0xFF81231Au;
    viewerTheme.alertWarningBackgroundArgb = 0xFFFFF0D0u;
    viewerTheme.alertWarningTextArgb       = 0xFF805000u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFFE0F0FFu;
    viewerTheme.alertInfoTextArgb          = 0xFF1A4A77u;
    viewerTheme.darkMode                   = FALSE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = FALSE;

    const ThemePalette theme           = MakeThemePaletteFromViewerTheme(viewerTheme);
    const TreeSurfaceVisualStyle style = ResolveTreeSurfaceVisualStyle(theme);

    RequireColorNear(style.fill, theme.surfaceBackground, "viewer-derived tree surface style uses the palette surface background");
    RequireColorNear(style.border, theme.border, "viewer-derived tree surface style uses the palette border");
    RequireColorNear(style.emptyText, theme.subduedText, "viewer-derived tree surface style empty text uses the palette subdued text");
    RequireColorDifferent(style.fill, theme.accent, "viewer-derived tree surface style follows tree chrome instead of raw accent");
}

void TestTreeSurfaceHighContrastUsesPaletteTreeChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const TreeSurfaceVisualStyle style = ResolveTreeSurfaceVisualStyle(theme);

    RequireColorNear(style.fill, theme.surfaceBackground, "high-contrast tree surface style keeps the palette surface background");
    RequireColorNear(style.border, theme.border, "high-contrast tree surface style keeps the palette border");
    RequireColorNear(style.emptyText, theme.subduedText, "high-contrast tree surface style keeps the palette subdued text");
}

void TestGridHeaderVisualStyleUsesSharedHeaderChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme          = MakeDefaultThemePalette(true);
    const GridHeaderVisualStyle style = ResolveGridHeaderVisualStyle(theme);

    RequireColorNear(style.fill, theme.headerBackground, "grid header style uses the palette header background");
    RequireColorNear(style.hoveredFill, theme.headerHovered, "grid header style uses the palette hovered header fill");
    RequireColorNear(style.pressedFill, theme.headerPressed, "grid header style uses the palette pressed header fill");
    RequireColorNear(style.separator, theme.gridLine, "grid header style uses the palette grid line for separators");
    RequireColorNear(style.titleText, theme.text, "grid header style uses the palette text for titles");
    RequireColorNear(style.busyGlyph, theme.selectionFill, "grid header style uses the palette selection fill for the busy glyph");
    RequireColorNear(style.sortGlyph, theme.subduedText, "grid header style uses the palette subdued text for sort glyphs");
    RequireColorNear(style.groupFill,
                     BlendForTest(theme.headerBackground, theme.surfaceBackground, 0.42f),
                     "grid header style derives grouped header fill from the shared header and surface chrome");
    RequireColorNear(style.groupSeparator, theme.gridLine, "grid header style uses the palette grid line for grouped separators");
    RequireColorNear(style.groupText, theme.subduedText, "grid header style uses the palette subdued text for grouped titles");
    RequireColorNear(style.groupGlyph, theme.subduedText, "grid header style uses the palette subdued text for grouped disclosure glyphs");
}

void TestGridHeaderVisualStyleUsesViewerDerivedHeaderChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF12171Eu;
    viewerTheme.textArgb                   = 0xFFE6ECF8u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4C7FD0u;
    viewerTheme.selectionTextArgb          = 0xFFF8FBFFu;
    viewerTheme.accentArgb                 = 0xFFD26457u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5D1F26u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD9DDu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF5B430Eu;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE3A2u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF18324Au;
    viewerTheme.alertInfoTextArgb          = 0xFFD7E8FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme          = MakeThemePaletteFromViewerTheme(viewerTheme);
    const GridHeaderVisualStyle style = ResolveGridHeaderVisualStyle(theme);

    RequireColorNear(style.fill, theme.headerBackground, "viewer-derived grid header style uses the palette header background");
    RequireColorNear(style.hoveredFill, theme.headerHovered, "viewer-derived grid header style uses the palette hovered header fill");
    RequireColorNear(style.pressedFill, theme.headerPressed, "viewer-derived grid header style uses the palette pressed header fill");
    RequireColorNear(style.separator, theme.gridLine, "viewer-derived grid header style uses the palette grid line for separators");
    RequireColorNear(style.titleText, theme.text, "viewer-derived grid header style uses the palette text for titles");
    RequireColorNear(style.busyGlyph, theme.selectionFill, "viewer-derived grid header style uses the palette selection fill for the busy glyph");
    RequireColorNear(style.sortGlyph, theme.subduedText, "viewer-derived grid header style uses the palette subdued text for sort glyphs");
    RequireColorNear(style.groupFill,
                     BlendForTest(theme.headerBackground, theme.surfaceBackground, 0.42f),
                     "viewer-derived grid header style derives grouped header fill from the shared header and surface chrome");
    RequireColorNear(style.groupText, theme.subduedText, "viewer-derived grid header style uses the palette subdued text for grouped titles");
    RequireColorDifferent(style.groupFill, theme.accent, "viewer-derived grouped header fill follows shared header chrome instead of raw accent");
    RequireColorDifferent(style.busyGlyph, theme.accent, "viewer-derived grid header busy glyph does not read raw accent");
}

void TestGridHeaderHighContrastUsesPaletteHeaderChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const GridHeaderVisualStyle style = ResolveGridHeaderVisualStyle(theme);

    RequireColorNear(style.fill, theme.headerBackground, "high-contrast grid header style keeps the palette header background");
    RequireColorNear(style.hoveredFill, theme.headerHovered, "high-contrast grid header style keeps the palette hovered header fill");
    RequireColorNear(style.pressedFill, theme.headerPressed, "high-contrast grid header style keeps the palette pressed header fill");
    RequireColorNear(style.separator, theme.gridLine, "high-contrast grid header style keeps the palette grid line for separators");
    RequireColorNear(style.titleText, theme.text, "high-contrast grid header style keeps the palette title text");
    RequireColorNear(style.busyGlyph, theme.selectionFill, "high-contrast grid header style keeps the palette busy glyph chrome");
    RequireColorNear(style.sortGlyph, theme.subduedText, "high-contrast grid header style keeps the palette sort glyph chrome");
    RequireColorNear(style.groupFill,
                     BlendForTest(theme.headerBackground, theme.surfaceBackground, 0.42f),
                     "high-contrast grid header style keeps grouped header fill derived from shared header and surface chrome");
    RequireColorNear(style.groupSeparator, theme.gridLine, "high-contrast grid header style keeps grouped separator chrome");
    RequireColorNear(style.groupText, theme.subduedText, "high-contrast grid header style keeps grouped title chrome");
    RequireColorNear(style.groupGlyph, theme.subduedText, "high-contrast grid header style keeps grouped glyph chrome");
}

void TestTooltipVisualStyleUsesSharedTooltipChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme       = MakeDefaultThemePalette(true);
    const TooltipVisualStyle style = ResolveTooltipVisualStyle(theme);

    RequireColorNear(style.fill, theme.tooltipBackground, "tooltip style uses the palette tooltip background");
    RequireColorNear(style.text, theme.tooltipText, "tooltip style uses the palette tooltip text");
    RequireColorNear(style.border,
                     BlendForTest(theme.tooltipBackground, theme.tooltipText, theme.highContrast ? 0.18f : 0.10f),
                     "tooltip style border derives from the shared tooltip chrome contract");
}

void TestTooltipUsesViewerDerivedTooltipChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme       = MakeThemePaletteFromViewerTheme(viewerTheme);
    const TooltipVisualStyle style = ResolveTooltipVisualStyle(theme);

    RequireColorNear(style.fill, theme.tooltipBackground, "viewer-derived tooltip style uses the palette tooltip background");
    RequireColorNear(style.text, theme.tooltipText, "viewer-derived tooltip style uses the palette tooltip text");
    RequireColorNear(style.border,
                     BlendForTest(theme.tooltipBackground, theme.tooltipText, theme.highContrast ? 0.18f : 0.10f),
                     "viewer-derived tooltip border derives from the shared tooltip chrome contract");
    RequireColorDifferent(style.border, theme.accent, "viewer-derived tooltip border follows tooltip chrome instead of raw accent");
}

void TestTooltipHighContrastBorderStrengthensLegibility()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette normalTheme = MakeDefaultThemePalette(false);
    ThemePalette highContrastTheme = normalTheme;
    highContrastTheme.highContrast = true;

    const TooltipVisualStyle normalStyle       = ResolveTooltipVisualStyle(normalTheme);
    const TooltipVisualStyle highContrastStyle = ResolveTooltipVisualStyle(highContrastTheme);

    RequireColorNear(highContrastStyle.fill, highContrastTheme.tooltipBackground, "high-contrast tooltip keeps the palette tooltip background");
    RequireColorNear(highContrastStyle.text, highContrastTheme.tooltipText, "high-contrast tooltip keeps the palette tooltip text");
    RequireColorNear(highContrastStyle.border,
                     BlendForTest(highContrastTheme.tooltipBackground, highContrastTheme.tooltipText, 0.18f),
                     "high-contrast tooltip border uses the stronger shared tooltip chrome blend");
    RequireColorDifferent(highContrastStyle.border, normalStyle.border, "high-contrast tooltip border stays distinct from the normal tooltip border");
}

void TestCheckboxVisualStyleMatchesStandardCheckboxChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme = MakeDefaultThemePalette(true);

    const CheckboxVisualStyle idle = ResolveCheckboxVisualStyle(theme, true, false, false, false, false, false);
    Require(! idle.showHoverFill, "idle checkbox keeps quiet background");
    RequireColorNear(idle.indicatorFill, theme.inputFill, "idle checkbox uses neutral indicator fill");

    const CheckboxVisualStyle hovered = ResolveCheckboxVisualStyle(theme, true, true, false, false, false, false);
    Require(hovered.showHoverFill, "hovered checkbox shows subtle hover fill");
    RequireColorNear(
        hovered.hoverFill,
        BlendForTest(theme.windowBackground, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.08f : 0.05f),
        "hovered checkbox uses the shared hover overlay token");

    const CheckboxVisualStyle pressed = ResolveCheckboxVisualStyle(theme, true, false, true, false, false, false);
    Require(pressed.showHoverFill, "pressed checkbox keeps hover fill visible");
    RequireColorNear(
        pressed.hoverFill,
        BlendForTest(theme.windowBackground, D2D1::ColorF(theme.pressedFill.r, theme.pressedFill.g, theme.pressedFill.b, 1.0f), theme.dark ? 0.14f : 0.09f),
        "pressed checkbox uses the shared pressed overlay token");

    const CheckboxVisualStyle checked = ResolveCheckboxVisualStyle(theme, true, false, false, false, false, true);
    RequireColorNear(checked.indicatorFill, theme.selectionFill, "checked checkbox uses shared selection fill for indicator chrome");
    RequireColorNear(checked.indicatorBorder,
                     BlendForTest(theme.selectionFill, theme.selectionText, theme.highContrast ? 0.24f : 0.08f),
                     "checked checkbox uses selection-fill and selection-text border chrome");
    RequireColorNear(checked.check, theme.selectionText, "checked checkbox uses palette-derived selection text chrome");

    const CheckboxVisualStyle pointerFocused = ResolveCheckboxVisualStyle(theme, true, false, false, true, false, false);
    Require(! pointerFocused.showFocus, "pointer-focused checkbox keeps focus chrome quiet");

    const CheckboxVisualStyle keyboardFocused = ResolveCheckboxVisualStyle(theme, true, false, false, true, true, false);
    Require(keyboardFocused.showFocus, "keyboard-focused checkbox shows focus chrome");

    ThemePalette highContrastTheme                = theme;
    highContrastTheme.highContrast                = true;
    const CheckboxVisualStyle highContrastFocused = ResolveCheckboxVisualStyle(highContrastTheme, true, false, false, true, false, false);
    Require(highContrastFocused.showFocus, "high-contrast checkbox keeps focus chrome visible even without keyboard focus visibility");

    const CheckboxVisualStyle highContrastChecked = ResolveCheckboxVisualStyle(highContrastTheme, true, false, false, false, false, true);
    RequireColorNear(highContrastChecked.indicatorBorder,
                     BlendForTest(highContrastTheme.selectionFill, highContrastTheme.selectionText, 0.24f),
                     "high-contrast checked checkbox uses the stronger selection border chrome blend");
    RequireColorDifferent(highContrastChecked.indicatorBorder,
                          checked.indicatorBorder,
                          "high-contrast checked checkbox keeps indicator border distinct from the normal-theme checked border");

    const CheckboxVisualStyle disabledChecked = ResolveCheckboxVisualStyle(theme, false, false, false, false, false, true);
    RequireColorNear(disabledChecked.text, theme.disabledText, "disabled checkbox uses disabled text color");
    RequireColorNear(disabledChecked.check, theme.selectionText, "disabled checked checkbox keeps palette-derived check chrome");
}

void TestLabelVisualStyleUsesPaletteTextChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme            = MakeDefaultThemePalette(true);
    const LabelVisualStyle defaultStyle = ResolveLabelVisualStyle(theme, std::nullopt);
    const D2D1_COLOR_F customText       = D2D1::ColorF(0.82f, 0.64f, 0.31f, 1.0f);
    const LabelVisualStyle customStyle  = ResolveLabelVisualStyle(theme, customText);

    RequireColorNear(defaultStyle.text, theme.text, "default label style uses the palette text color");
    RequireColorNear(customStyle.text, customText, "label style honors explicit text color overrides");
}

void TestLabelVisualStyleUsesViewerDerivedTextChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme            = MakeThemePaletteFromViewerTheme(viewerTheme);
    const LabelVisualStyle defaultStyle = ResolveLabelVisualStyle(theme, std::nullopt);
    const D2D1_COLOR_F customText       = D2D1::ColorF(0.91f, 0.80f, 0.44f, 1.0f);
    const LabelVisualStyle customStyle  = ResolveLabelVisualStyle(theme, customText);

    RequireColorNear(defaultStyle.text, theme.text, "viewer-derived label style uses the palette text color");
    RequireColorNear(customStyle.text, customText, "viewer-derived label style honors explicit text color overrides");
}

void TestLabelVisualStyleUsesHighContrastTextChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const D2D1_COLOR_F customText       = D2D1::ColorF(0.18f, 0.42f, 0.82f, 1.0f);
    const LabelVisualStyle defaultStyle = ResolveLabelVisualStyle(theme, std::nullopt);
    const LabelVisualStyle customStyle  = ResolveLabelVisualStyle(theme, customText);

    RequireColorNear(defaultStyle.text, theme.text, "high-contrast label style keeps the palette text color");
    RequireColorNear(customStyle.text, customText, "high-contrast label style still honors explicit text color overrides");
}

void TestCheckboxVisualStyleInterpolatesHoverStrength()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme            = MakeDefaultThemePalette(true);
    const CheckboxVisualStyle halfHover = ResolveCheckboxVisualStyle(theme, true, true, false, false, false, false, 0.5f, 0.0f);
    const D2D1_COLOR_F expectedHoverFill =
        BlendForTest(theme.windowBackground, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), (theme.dark ? 0.08f : 0.05f) * 0.5f);
    Require(halfHover.showHoverFill, "checkbox hover animation keeps hover fill visible while in flight");
    RequireColorNear(halfHover.hoverFill, expectedHoverFill, "checkbox hover animation blends hover fill by progress");
}

void TestCheckboxVisualStyleUsesViewerDerivedSelectionTextChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF2D6FB7u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme          = MakeThemePaletteFromViewerTheme(viewerTheme);
    const CheckboxVisualStyle idle    = ResolveCheckboxVisualStyle(theme, true, false, false, false, false, false);
    const CheckboxVisualStyle hovered = ResolveCheckboxVisualStyle(theme, true, true, false, false, false, false);
    const CheckboxVisualStyle pressed = ResolveCheckboxVisualStyle(theme, true, false, true, false, false, false);
    const CheckboxVisualStyle checked = ResolveCheckboxVisualStyle(theme, true, false, false, false, false, true);

    RequireColorNear(idle.indicatorBorder, theme.inputBorder, "viewer-derived checkbox idle border chrome uses the palette input border");
    RequireColorNear(idle.check, theme.text, "viewer-derived checkbox idle chrome uses the palette text color");
    RequireColorNear(
        hovered.hoverFill,
        BlendForTest(theme.windowBackground, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.08f : 0.05f),
        "viewer-derived checkbox hover fill uses the shared hover overlay token");
    RequireColorNear(
        pressed.hoverFill,
        BlendForTest(theme.windowBackground, D2D1::ColorF(theme.pressedFill.r, theme.pressedFill.g, theme.pressedFill.b, 1.0f), theme.dark ? 0.14f : 0.09f),
        "viewer-derived checkbox pressed fill uses the shared pressed overlay token");
    RequireColorNear(checked.indicatorFill, theme.selectionFill, "viewer-derived checkbox checked fill uses the shared selection fill");
    RequireColorNear(checked.indicatorBorder,
                     BlendForTest(theme.selectionFill, theme.selectionText, theme.highContrast ? 0.24f : 0.08f),
                     "viewer-derived checkbox checked border chrome uses shared selection fill and text");
    RequireColorNear(checked.check, theme.selectionText, "viewer-derived checkbox checked chrome uses the palette selection text");
    RequireColorDifferent(checked.indicatorFill, theme.accent, "viewer-derived checkbox checked fill follows selection fill instead of raw accent");
    RequireColorDifferent(idle.check, checked.check, "viewer-derived checkbox keeps idle and checked glyph chrome distinct");
}

void TestCheckboxHighContrastDisabledBordersStayVisible()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(true);
    theme.highContrast = true;

    const CheckboxVisualStyle disabledUnchecked = ResolveCheckboxVisualStyle(theme, false, false, false, false, false, false);
    const CheckboxVisualStyle disabledChecked   = ResolveCheckboxVisualStyle(theme, false, false, false, false, false, true);

    const D2D1_COLOR_F expectedDisabledFill   = BlendForTest(theme.windowBackground, theme.inputFill, theme.dark ? 0.72f : 0.58f);
    const D2D1_COLOR_F expectedDisabledBorder = BlendForTest(theme.windowBackground, theme.inputBorder, theme.dark ? 0.60f : 0.44f);

    RequireColorNear(
        disabledUnchecked.indicatorFill, expectedDisabledFill, "high-contrast disabled unchecked checkbox keeps the shared disabled indicator fill");
    RequireColorNear(
        disabledUnchecked.indicatorBorder, expectedDisabledBorder, "high-contrast disabled unchecked checkbox keeps the shared disabled indicator border");
    RequireColorNear(disabledUnchecked.text, theme.disabledText, "high-contrast disabled unchecked checkbox keeps disabled text color");

    RequireColorNear(disabledChecked.indicatorFill,
                     BlendForTest(expectedDisabledFill, theme.selectionFill, theme.dark ? 0.28f : 0.18f),
                     "high-contrast disabled checked checkbox keeps the shared disabled-checked indicator fill blend");
    RequireColorNear(disabledChecked.indicatorBorder,
                     expectedDisabledBorder,
                     "high-contrast disabled checked checkbox keeps the shared disabled indicator border instead of dropping border chrome");
    RequireColorNear(disabledChecked.check, theme.selectionText, "high-contrast disabled checked checkbox keeps selection-text glyph chrome");
    RequireColorNear(disabledChecked.text, theme.disabledText, "high-contrast disabled checked checkbox keeps disabled text color");
    RequireColorDifferent(
        disabledChecked.indicatorFill, disabledUnchecked.indicatorFill, "high-contrast disabled checked checkbox keeps distinct checked indicator fill chrome");
}

void TestComboBoxVariantsExposeDistinctChromeContracts()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme = MakeDefaultThemePalette(true);

    const ComboBoxVisualStyle windowStyle = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Window, true, false, false, false, false);
    Require(windowStyle.showButtonSplit, "window combo shows split button");
    Require(! windowStyle.showLeftFocusAccent, "window combo does not show xaml focus bar");
    RequireColorNear(windowStyle.buttonBorder, theme.buttonBorder, "window combo idle button uses the palette button border");
    RequireColorNear(windowStyle.splitStroke, theme.inputBorder, "window combo split divider uses the palette input border");
    RequireColorNear(windowStyle.selectionFill, theme.selectionInactiveFill, "window combo idle selection uses the inactive selection fill");
    RequireColorNear(windowStyle.selectionText, theme.selectionText, "window combo idle selection text uses the palette selection text");
    RequireColorNear(windowStyle.caret, theme.focusStroke, "window combo caret uses the palette focus stroke");
    RequireColorNear(windowStyle.glyph, theme.text, "window combo glyph uses the palette text chrome");
    RequireColorNear(windowStyle.popupFill, theme.overlayBackground, "window combo popup uses the shared overlay background");
    RequireColorNear(windowStyle.popupBorder, theme.overlayBorder, "window combo popup uses the shared overlay border");
    RequireColorNear(windowStyle.popupSelectedFill, theme.selectionInactiveFill, "window combo selected popup rows use the inactive selection fill");
    RequireColorNear(windowStyle.popupActiveFill, theme.selectionFill, "window combo active popup rows use the active selection fill");
    RequireColorNear(windowStyle.popupText, theme.text, "window combo popup rows use the palette text");
    RequireColorNear(windowStyle.popupSelectedText, theme.text, "window combo selected popup rows keep the regular body text color");
    RequireColorNear(windowStyle.popupActiveText, theme.selectionText, "window combo active popup rows use the palette selection text");
    RequireColorNear(windowStyle.popupEmptyText, theme.subduedText, "window combo popup empty state uses the palette subdued text");
    RequireColorNear(windowStyle.popupScrollbarTrack, theme.scrollbarTrack, "window combo popup scrollbar track uses the palette track color");
    RequireColorNear(windowStyle.popupScrollbarThumb, theme.scrollbarThumb, "window combo popup scrollbar thumb uses the palette thumb color");
    RequireColorNear(windowStyle.popupScrollbarThumbHot, theme.scrollbarThumbHot, "window combo popup scrollbar hot thumb uses the palette hot thumb color");

    const ComboBoxVisualStyle modernStyle = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Modern, true, false, false, true, true);
    Require(modernStyle.showLeftFocusAccent, "modern combo shows left focus accent");
    Require(! modernStyle.showButtonSplit, "modern combo hides classic split button");
    RequireColorNear(modernStyle.selectionFill, theme.selectionFill, "focused modern combo uses the active selection fill");
    RequireColorNear(modernStyle.selectionText, theme.selectionText, "focused modern combo uses the palette selection text");
    RequireColorNear(modernStyle.caret, theme.focusStroke, "focused modern combo caret uses the palette focus stroke");
    RequireColorNear(modernStyle.glyph, theme.text, "focused modern combo glyph uses the palette text chrome");

    const ComboBoxVisualStyle modernHovered = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Modern, true, true, false, false, false);
    Require(modernHovered.showButtonBackground, "modern combo shows button background while hovered");
    RequireColorNear(modernHovered.buttonFill, theme.buttonHotFill, "modern combo hovered button uses the palette hot fill");

    const ComboBoxVisualStyle modernPopupOpen = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Modern, true, false, true, true, true);
    Require(modernPopupOpen.showButtonBackground, "modern combo shows button background while popup is open");
    RequireColorNear(modernPopupOpen.buttonFill, theme.buttonPressedFill, "modern combo popup-open button uses the palette pressed fill");

    const ComboBoxVisualStyle editStyle = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Edit, true, false, false, true, true);
    Require(editStyle.showButtonSplit, "edit combo keeps split button");
    Require(! editStyle.showLeftFocusAccent, "edit combo does not use xaml focus bar");
    RequireColorNear(editStyle.splitStroke, editStyle.fieldBorder, "edit combo keeps split divider chrome aligned with the resolved field border");

    ComboBox combo;
    Require(combo.GetVariant() == ComboBoxVariant::Modern, "combo defaults to modern variant");
    Require(! combo.IsEditable(), "default combo starts non-editable");

    combo.SetVariant(ComboBoxVariant::Window);
    Require(combo.GetVariant() == ComboBoxVariant::Window, "combo stores explicit window variant");
    Require(! combo.IsEditable(), "window combo remains non-editable");

    combo.SetEditable(true);
    Require(combo.GetVariant() == ComboBoxVariant::Edit, "editable combo switches to edit variant");
    Require(combo.IsEditable(), "edit variant is editable");

    combo.SetEditable(false);
    Require(combo.GetVariant() == ComboBoxVariant::Modern, "disabling edit mode returns combo to modern variant");
    Require(! combo.IsEditable(), "combo stops being editable when edit mode is cleared");
}

void TestComboBoxHoverStyleStrengthensFieldChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme          = MakeDefaultThemePalette(true);
    const ComboBoxVisualStyle idle    = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Window, true, false, false, false, false);
    const ComboBoxVisualStyle hovered = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Window, true, true, false, false, false);

    RequireColorDifferent(hovered.fieldBorder, idle.fieldBorder, "hovered combo field border differs from idle");
    RequireColorNear(hovered.fieldFill,
                     BlendForTest(idle.fieldFill, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.05f : 0.025f),
                     "hovered combo field fill uses the shared hover overlay token");
    RequireColorNear(hovered.buttonBorder,
                     BlendForTest(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.16f : 0.12f),
                     "hovered window combo button derives border chrome from the palette button border");
}

void TestComboBoxModernPopupUsesViewerDerivedPressedChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF4F8EDCu;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme                  = MakeThemePaletteFromViewerTheme(viewerTheme);
    const ComboBoxVisualStyle modernHovered   = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Modern, true, true, false, false, false);
    const ComboBoxVisualStyle modernPopupOpen = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Modern, true, false, true, true, true);

    Require(modernHovered.showButtonBackground, "viewer-derived modern combo shows button background while hovered");
    RequireColorNear(modernHovered.buttonFill, theme.buttonHotFill, "viewer-derived modern combo hovered button uses the palette hot fill");
    Require(modernPopupOpen.showButtonBackground, "viewer-derived modern combo shows button background while popup is open");
    RequireColorNear(modernPopupOpen.buttonFill, theme.buttonPressedFill, "viewer-derived modern combo popup-open button uses the palette pressed fill");
}

void TestComboBoxUsesViewerDerivedInputBorderChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF4F8EDCu;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme              = MakeThemePaletteFromViewerTheme(viewerTheme);
    const ComboBoxVisualStyle idle        = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Window, true, false, false, false, false);
    const ComboBoxVisualStyle hovered     = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Window, true, true, false, false, false);
    const ComboBoxVisualStyle focused     = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Modern, true, false, false, true, true);
    const ComboBoxVisualStyle editFocused = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Edit, true, false, false, true, true);

    RequireColorNear(idle.fieldBorder, theme.inputBorder, "viewer-derived combo idle field border uses the palette input border");
    RequireColorNear(idle.splitStroke, theme.inputBorder, "viewer-derived combo split divider uses the palette input border");
    RequireColorNear(idle.selectionFill, theme.selectionInactiveFill, "viewer-derived idle combo selection uses the inactive selection fill");
    RequireColorNear(idle.selectionText, theme.selectionText, "viewer-derived idle combo selection text uses the palette selection text");
    RequireColorNear(idle.caret, theme.focusStroke, "viewer-derived idle combo caret uses the palette focus stroke");
    RequireColorNear(idle.glyph, theme.text, "viewer-derived idle combo glyph uses the palette text chrome");
    RequireColorNear(idle.popupFill, theme.overlayBackground, "viewer-derived combo popup uses the shared overlay background");
    RequireColorNear(idle.popupBorder, theme.overlayBorder, "viewer-derived combo popup uses the shared overlay border");
    RequireColorNear(idle.popupSelectedFill, theme.selectionInactiveFill, "viewer-derived combo selected popup rows use the inactive selection fill");
    RequireColorNear(idle.popupActiveFill, theme.selectionFill, "viewer-derived combo active popup rows use the active selection fill");
    RequireColorNear(idle.popupText, theme.text, "viewer-derived combo popup rows use the palette text");
    RequireColorNear(idle.popupSelectedText, theme.text, "viewer-derived combo selected popup rows keep the regular body text color");
    RequireColorNear(idle.popupActiveText, theme.selectionText, "viewer-derived combo active popup rows use the palette selection text");
    RequireColorNear(idle.popupEmptyText, theme.subduedText, "viewer-derived combo popup empty state uses the palette subdued text");
    RequireColorNear(idle.popupScrollbarTrack, theme.scrollbarTrack, "viewer-derived combo popup scrollbar track uses the palette track color");
    RequireColorNear(idle.popupScrollbarThumb, theme.scrollbarThumb, "viewer-derived combo popup scrollbar thumb uses the palette thumb color");
    RequireColorNear(idle.popupScrollbarThumbHot, theme.scrollbarThumbHot, "viewer-derived combo popup scrollbar hot thumb uses the palette hot thumb color");
    RequireColorNear(hovered.fieldFill,
                     BlendForTest(idle.fieldFill, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.05f : 0.025f),
                     "viewer-derived combo hover field fill uses the shared hover overlay token");
    RequireColorNear(focused.fieldBorder,
                     BlendForTest(theme.inputBorder, theme.focusStroke, theme.dark ? 0.60f : 0.72f),
                     "viewer-derived modern combo focus border uses the palette input border");
    RequireColorNear(
        editFocused.splitStroke, editFocused.fieldBorder, "viewer-derived edit combo keeps split divider chrome aligned with the resolved field border");
    RequireColorNear(focused.selectionFill, theme.selectionFill, "viewer-derived focused combo selection uses the active selection fill");
    RequireColorNear(focused.selectionText, theme.selectionText, "viewer-derived focused combo selection text uses the palette selection text");
    RequireColorNear(focused.caret, theme.focusStroke, "viewer-derived focused combo caret uses the palette focus stroke");
    RequireColorNear(focused.glyph, theme.text, "viewer-derived focused combo glyph uses the palette text chrome");
}

void TestComboBoxHighContrastFocusAccentStaysVisibleWithoutKeyboardFocus()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const ComboBoxVisualStyle modernPointerFocused = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Modern, true, false, false, true, false);
    const ComboBoxVisualStyle editPointerFocused   = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Edit, true, false, false, true, false);

    Require(modernPointerFocused.showLeftFocusAccent, "high-contrast modern combo keeps left focus accent visible without keyboard-focus gating");
    RequireColorNear(modernPointerFocused.focusAccent, theme.focusStroke, "high-contrast modern combo focus accent uses the palette focus stroke");
    Require(! editPointerFocused.showLeftFocusAccent, "high-contrast edit combo still keeps the text-field split layout instead of a left focus accent");
}

void TestComboBoxHighContrastDisabledBordersStayVisible()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const ComboBoxVisualStyle disabledWindow = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Window, false, false, false, false, false);
    const ComboBoxVisualStyle disabledEdit   = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Edit, false, false, false, false, false);

    RequireColorNear(disabledWindow.fieldBorder,
                     BlendForTest(theme.windowBackground, theme.inputBorder, theme.dark ? 0.60f : 0.44f),
                     "high-contrast disabled window combo keeps the shared disabled input border chrome");
    RequireColorNear(disabledWindow.buttonBorder,
                     BlendForTest(theme.windowBackground, theme.buttonBorder, theme.dark ? 0.60f : 0.44f),
                     "high-contrast disabled window combo keeps the shared disabled button border chrome");
    RequireColorNear(disabledWindow.text, theme.disabledText, "high-contrast disabled window combo keeps disabled text color");
    RequireColorNear(disabledWindow.placeholderText, theme.disabledText, "high-contrast disabled window combo keeps disabled placeholder text color");

    RequireColorNear(disabledEdit.fieldBorder,
                     BlendForTest(theme.windowBackground, theme.inputBorder, theme.dark ? 0.60f : 0.44f),
                     "high-contrast disabled edit combo keeps the shared disabled input border chrome");
    RequireColorNear(
        disabledEdit.buttonBorder, disabledEdit.fieldBorder, "high-contrast disabled edit combo keeps the split button border aligned with the field border");
    RequireColorNear(
        disabledEdit.splitStroke, disabledEdit.fieldBorder, "high-contrast disabled edit combo keeps the split stroke aligned with the field border");
}

void TestGridRowIconChromeFollowsResolvedRowVisuals()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    MultiRowGridModel model(3u);

    grid.SetModel(&model);
    host.SetFocusControl(&grid);

    const ThemePalette theme = MakeDefaultThemePalette(true);
    GridDebugRowVisualState state{};

    Require(grid.DebugGetRowVisualState(theme, 0u, state), "grid row icon visual test resolves the unselected row state");
    Require(! state.selected, "grid row icon visual test keeps the first row unselected");
    Require(state.iconArgb == PackColorForTest(ResolveListIconColor(theme, theme.text, false)), "grid unselected row uses the shared list-icon chrome");
    Require(state.busyArgb == PackColorForTest(theme.selectionFill), "grid unselected row exposes busy chrome through the shared row-visual contract");
    Require(state.progressTrackArgb != 0u, "grid unselected row exposes progress track chrome through the shared row-visual contract");
    Require(state.progressFillArgb == PackColorForTest(theme.selectionFill),
            "grid unselected row exposes progress fill chrome through the shared row-visual contract");

    grid.GetSelectionModel().SetSingle(model.GetStableRowId(1u));
    Require(grid.DebugGetRowVisualState(theme, 1u, state), "grid row icon visual test resolves the selected row state");
    Require(state.selected, "grid row icon visual test keeps the selected row selected");
    Require(state.iconArgb == state.textArgb, "grid selected row keeps icon chrome aligned with selected text chrome");
    Require(state.busyArgb == state.textArgb, "grid selected row keeps busy chrome aligned with selected text chrome");
    Require(state.progressTrackArgb != 0u, "grid selected row exposes progress track chrome through the shared row-visual contract");
    Require(state.progressFillArgb == state.textArgb, "grid selected row keeps progress fill chrome aligned with selected text chrome");
}

void TestGridRowIconChromeFollowsViewerDerivedRowVisuals()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    MultiRowGridModel model(3u);

    grid.SetModel(&model);
    host.SetFocusControl(&grid);

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF12171Eu;
    viewerTheme.textArgb                   = 0xFFE6ECF8u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4C7FD0u;
    viewerTheme.selectionTextArgb          = 0xFFF8FBFFu;
    viewerTheme.accentArgb                 = 0xFFD26457u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5D1F26u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD9DDu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF5B430Eu;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE3A2u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF18324Au;
    viewerTheme.alertInfoTextArgb          = 0xFFD7E8FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme = MakeThemePaletteFromViewerTheme(viewerTheme);
    GridDebugRowVisualState state{};

    Require(grid.DebugGetRowVisualState(theme, 0u, state), "grid row icon visual test resolves the viewer-derived unselected row state");
    Require(! state.selected, "grid viewer-derived unselected row stays unselected");
    Require(state.iconArgb == PackColorForTest(ResolveListIconColor(theme, theme.text, false)),
            "grid viewer-derived unselected row uses the shared list-icon chrome");
    Require(state.iconArgb != PackColorForTest(theme.accent), "grid viewer-derived unselected row icon chrome no longer follows raw accent");
    Require(state.busyArgb == PackColorForTest(theme.selectionFill),
            "grid viewer-derived unselected row exposes busy chrome through the shared row-visual contract");
    Require(state.busyArgb != PackColorForTest(theme.accent), "grid viewer-derived unselected row busy chrome does not fall back to raw accent");
    Require(state.progressTrackArgb != 0u, "grid viewer-derived unselected row exposes progress track chrome through the shared row-visual contract");
    Require(state.progressFillArgb == PackColorForTest(theme.selectionFill),
            "grid viewer-derived unselected row exposes progress fill chrome through the shared row-visual contract");
    Require(state.progressFillArgb != PackColorForTest(theme.accent), "grid viewer-derived unselected row progress fill does not fall back to raw accent");

    grid.GetSelectionModel().SetSingle(model.GetStableRowId(1u));
    Require(grid.DebugGetRowVisualState(theme, 1u, state), "grid row icon visual test resolves the viewer-derived selected row state");
    Require(state.selected, "grid viewer-derived selected row stays selected");
    Require(state.iconArgb == state.textArgb, "grid viewer-derived selected row keeps icon chrome aligned with selected text chrome");
    Require(state.busyArgb == state.textArgb, "grid viewer-derived selected row keeps busy chrome aligned with selected text chrome");
    Require(state.progressTrackArgb != 0u, "grid viewer-derived selected row exposes progress track chrome through the shared row-visual contract");
    Require(state.progressFillArgb == state.textArgb, "grid viewer-derived selected row keeps progress fill chrome aligned with selected text chrome");
}

void TestGridCellChromeFollowsResolvedCellVisuals()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme    = MakeDefaultThemePalette(true);
    const auto requireCellState = [&](const GridCellData& cellData, bool selected, const auto& verifier)
    {
        WindowHost host;
        Grid grid;
        SingleCellGridModel model(cellData);
        grid.SetModel(&model);
        host.SetFocusControl(&grid);
        if (selected)
        {
            grid.GetSelectionModel().SetSingle(model.GetStableRowId(0u));
        }

        GridDebugRowVisualState rowState{};
        GridDebugCellVisualState cellState{};
        Require(grid.DebugGetRowVisualState(theme, 0u, rowState), "grid cell visual test resolves row visuals");
        Require(grid.DebugGetCellVisualState(theme, 0u, 0u, cellState), "grid cell visual test resolves cell visuals");
        Require(cellState.selected == selected, "grid cell visual test preserves selected state");
        verifier(rowState, cellState);
    };

    GridCellData checkboxCell{};
    checkboxCell.kind    = GridCellKind::Checkbox;
    checkboxCell.checked = true;
    requireCellState(checkboxCell,
                     false,
                     [&](const GridDebugRowVisualState&, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasCheckbox, "grid checkbox cell exposes checkbox chrome");
        Require(cellState.checkboxCheckArgb == PackColorForTest(theme.selectionText), "grid unselected checkbox cell keeps the shared checkbox check chrome");
    });
    requireCellState(checkboxCell,
                     true,
                     [&](const GridDebugRowVisualState& rowState, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasCheckbox, "grid selected checkbox cell exposes checkbox chrome");
        Require(cellState.checkboxIndicatorFillArgb == rowState.textArgb, "grid selected checkbox cell inverts indicator fill to selected row text");
        Require(cellState.checkboxIndicatorBorderArgb == rowState.textArgb, "grid selected checkbox cell inverts indicator border to selected row text");
        Require(cellState.checkboxCheckArgb == rowState.fillArgb, "grid selected checkbox cell inverts check chrome to selected row fill");
    });

    GridCellData swatchCell{};
    swatchCell.kind           = GridCellKind::ColorSwatch;
    swatchCell.hasSwatchValue = true;
    swatchCell.swatchArgb     = 0x8044AA33u;
    requireCellState(swatchCell,
                     true,
                     [&](const GridDebugRowVisualState&, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasSwatch, "grid swatch cell exposes swatch chrome");
        Require(cellState.swatchBorderArgb ==
                    PackColorForTest(ResolveGridSwatchVisualStyle(theme, theme.selectionFill, theme.selectionText, true, swatchCell).border),
                "grid selected swatch cell keeps border chrome aligned with the shared selected swatch contract");
    });

    GridCellData badgeCell{};
    badgeCell.badgeText = L"Beta";
    badgeCell.badgeTone = AdornmentTone::Warning;
    requireCellState(badgeCell,
                     false,
                     [&](const GridDebugRowVisualState&, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasBadge, "grid badge cell exposes badge chrome");
        Require(cellState.badgeFillArgb == PackColorForTest(theme.warningFill), "grid unselected badge cell keeps the shared warning fill chrome");
    });
    requireCellState(badgeCell,
                     true,
                     [&](const GridDebugRowVisualState& rowState, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasBadge, "grid selected badge cell exposes badge chrome");
        Require(cellState.badgeFillArgb == rowState.textArgb, "grid selected badge cell inverts fill to selected row text");
        Require(cellState.badgeTextArgb == rowState.fillArgb, "grid selected badge cell inverts text chrome to selected row fill");
    });
}

void TestGridCellChromeFollowsViewerDerivedResolvedCellVisuals()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF12171Eu;
    viewerTheme.textArgb                   = 0xFFE6ECF8u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4C7FD0u;
    viewerTheme.selectionTextArgb          = 0xFFF8FBFFu;
    viewerTheme.accentArgb                 = 0xFFD26457u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5D1F26u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD9DDu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF5B430Eu;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE3A2u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF18324Au;
    viewerTheme.alertInfoTextArgb          = 0xFFD7E8FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme    = MakeThemePaletteFromViewerTheme(viewerTheme);
    const auto requireCellState = [&](const GridCellData& cellData, bool selected, const auto& verifier)
    {
        WindowHost host;
        Grid grid;
        SingleCellGridModel model(cellData);
        grid.SetModel(&model);
        host.SetFocusControl(&grid);
        if (selected)
        {
            grid.GetSelectionModel().SetSingle(model.GetStableRowId(0u));
        }

        GridDebugRowVisualState rowState{};
        GridDebugCellVisualState cellState{};
        Require(grid.DebugGetRowVisualState(theme, 0u, rowState), "viewer-derived grid cell visual test resolves row visuals");
        Require(grid.DebugGetCellVisualState(theme, 0u, 0u, cellState), "viewer-derived grid cell visual test resolves cell visuals");
        Require(cellState.selected == selected, "viewer-derived grid cell visual test preserves selected state");
        verifier(rowState, cellState);
    };

    GridCellData checkboxCell{};
    checkboxCell.kind    = GridCellKind::Checkbox;
    checkboxCell.checked = true;
    requireCellState(checkboxCell,
                     false,
                     [&](const GridDebugRowVisualState&, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasCheckbox, "viewer-derived grid checkbox cell exposes checkbox chrome");
        Require(cellState.checkboxCheckArgb == PackColorForTest(theme.selectionText),
                "viewer-derived grid unselected checkbox cell keeps the shared checkbox check chrome");
        Require(cellState.checkboxCheckArgb != PackColorForTest(theme.accent), "viewer-derived grid checkbox check chrome does not fall back to raw accent");
    });
    requireCellState(checkboxCell,
                     true,
                     [&](const GridDebugRowVisualState& rowState, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasCheckbox, "viewer-derived selected grid checkbox cell exposes checkbox chrome");
        Require(cellState.checkboxIndicatorFillArgb == rowState.textArgb,
                "viewer-derived selected grid checkbox cell inverts indicator fill to selected row text");
        Require(cellState.checkboxIndicatorBorderArgb == rowState.textArgb,
                "viewer-derived selected grid checkbox cell inverts indicator border to selected row text");
        Require(cellState.checkboxCheckArgb == rowState.fillArgb, "viewer-derived selected grid checkbox cell inverts check chrome to selected row fill");
        Require(cellState.checkboxIndicatorFillArgb != PackColorForTest(theme.accent),
                "viewer-derived selected grid checkbox indicator fill does not fall back to raw accent");
    });

    GridCellData swatchCell{};
    swatchCell.kind           = GridCellKind::ColorSwatch;
    swatchCell.hasSwatchValue = true;
    swatchCell.swatchArgb     = 0xCC55B6E8u;
    requireCellState(swatchCell,
                     true,
                     [&](const GridDebugRowVisualState&, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasSwatch, "viewer-derived grid swatch cell exposes swatch chrome");
        Require(cellState.swatchBorderArgb ==
                    PackColorForTest(ResolveGridSwatchVisualStyle(theme, theme.selectionFill, theme.selectionText, true, swatchCell).border),
                "viewer-derived grid selected swatch cell keeps border chrome aligned with the shared selected swatch contract");
        Require(cellState.swatchBorderArgb != PackColorForTest(theme.accent), "viewer-derived grid swatch border does not fall back to raw accent");
    });

    GridCellData badgeCell{};
    badgeCell.badgeText = L"Warn";
    badgeCell.badgeTone = AdornmentTone::Warning;
    requireCellState(badgeCell,
                     true,
                     [&](const GridDebugRowVisualState& rowState, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasBadge, "viewer-derived grid badge cell exposes badge chrome");
        Require(cellState.badgeFillArgb == rowState.textArgb, "viewer-derived selected grid badge cell inverts fill to selected row text");
        Require(cellState.badgeTextArgb == rowState.fillArgb, "viewer-derived selected grid badge cell inverts text chrome to selected row fill");
        Require(cellState.badgeFillArgb != PackColorForTest(theme.accent), "viewer-derived grid selected badge fill does not fall back to raw accent");
    });
}

// ---------------------------------------------------------------------------
// Phase 0-4 new token and visual style tests
// ---------------------------------------------------------------------------

void TestNewTokensPresentInDarkDefaultPalette()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette dark = MakeDefaultThemePalette(true);

    Require(dark.smokeOverlay.a > 0.0f, "dark default palette smoke overlay has non-zero alpha");
    RequireColorDifferent(dark.borderDefault, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "dark default palette border default is non-zero");
    RequireColorDifferent(dark.borderStrong, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "dark default palette border strong is non-zero");
    RequireColorDifferent(dark.accentHover, dark.accent, "dark default palette accent hover differs from accent");
    RequireColorDifferent(dark.accentPressed, dark.accent, "dark default palette accent pressed differs from accent");
    RequireColorDifferent(dark.focusStrokeOuter, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "dark default palette focus stroke outer is non-zero");
    RequireColorDifferent(dark.focusStrokeInner, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "dark default palette focus stroke inner is non-zero");
    RequireColorDifferent(dark.focusStrokeOuter, dark.focusStrokeInner, "dark default palette focus stroke outer and inner differ");
    RequireColorDifferent(dark.cardBackground, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "dark default palette card background is non-zero");
}

void TestDarkDefaultPaletteOverridesLightInteractionTokens()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette light = MakeDefaultThemePalette(false);
    const ThemePalette dark  = MakeDefaultThemePalette(true);

    RequireColorDifferent(dark.accent, light.accent, "dark default palette accent does not inherit the light accent token");
    RequireColorNear(
        dark.focusStroke, BlendForTest(dark.selectionFill, dark.selectionText, 0.10f), "dark default palette derives focus stroke from dark selection chrome");
    RequireColorDifferent(dark.focusStroke, light.focusStroke, "dark default palette focus stroke does not inherit the light focus token");
    RequireColorNear(dark.pressedFill,
                     D2D1::ColorF(dark.focusStroke.r, dark.focusStroke.g, dark.focusStroke.b, 0.24f),
                     "dark default palette derives pressed fill from the dark focus stroke");
    Require(dark.pressedFill.a > dark.hoverFill.a, "dark default palette pressed fill is stronger than hover fill");
    RequireColorNear(dark.toggleKnobFill,
                     ChooseContrastingTextColorForTest(BlendForTest(dark.inputFill, dark.border, 0.18f)),
                     "dark default palette derives neutral toggle knob chrome from dark neutral colors");
    RequireColorNear(dark.toggleKnobCheckedFill, dark.selectionText, "dark default palette checked toggle knob uses dark selection text");
    Require(dark.smokeOverlay.a > light.smokeOverlay.a, "dark default palette smoke overlay is stronger than the light overlay");
}

void TestAppThemeDxPaletteRefreshesAccentVariantsAfterAssigningAccent()
{
    const D2D1::ColorF customAccent(0.82f, 0.22f, 0.52f, 1.0f);
    const AppTheme theme                                   = ResolveAppTheme(ThemeMode::Rainbow, L"dx-palette-accent-refresh", customAccent);
    const RedSalamander::DxUi::ThemePalette palette        = MakeAppThemeDxPalette(theme);
    const RedSalamander::DxUi::ThemePalette defaultPalette = RedSalamander::DxUi::MakeDefaultThemePalette(theme.dark);

    RequireColorNear(palette.accent, customAccent, "app Dx palette keeps the app-assigned accent");
    RequireColorDifferent(palette.accentHover, defaultPalette.accentHover, "app Dx palette accent hover is not the stale default variant");
    RequireColorDifferent(palette.accentPressed, defaultPalette.accentPressed, "app Dx palette accent pressed is not the stale default variant");
    RequireColorDifferent(palette.headerHovered,
                          BlendForTest(palette.overlayBackground, defaultPalette.accentHover, theme.dark ? 0.44f : 0.28f),
                          "rainbow app Dx palette header hover uses the refreshed accent-hover variant");
    RequireColorDifferent(palette.headerPressed,
                          BlendForTest(palette.overlayBackground, defaultPalette.accentPressed, theme.dark ? 0.62f : 0.42f),
                          "rainbow app Dx palette header pressed uses the refreshed accent-pressed variant");
}

void TestFolderContentDxPalettePreservesItsNamedProfile()
{
    const AppTheme theme                            = ResolveAppTheme(ThemeMode::Rainbow, L"folder-content-palette-golden");
    const RedSalamander::DxUi::ThemePalette palette = MakeFolderContentDxPalette(theme);

    Require(palette.dark == theme.dark && palette.highContrast == theme.highContrast && palette.rainbowMode == theme.menu.rainbowMode,
            "folder-content palette preserves app mode flags");
    RequireColorNear(palette.accent, theme.accent, "folder-content palette preserves the resolved app accent");
    RequireColorNear(palette.surfaceBackground, theme.folderView.backgroundColor, "folder-content palette uses the folder-view surface");
    RequireColorNear(palette.inputFill, theme.folderView.backgroundColor, "folder-content palette keeps inputs on the folder-view surface");
    RequireColorNear(palette.gridLine, theme.folderView.gridLines, "folder-content palette uses folder-view grid lines");
    RequireColorNear(palette.text, theme.folderView.textNormal, "folder-content palette uses folder-view text");
    RequireColorNear(palette.headerPressed,
                     BlendForTest(palette.headerBackground, palette.accent, theme.dark ? 0.32f : 0.18f),
                     "folder-content palette preserves its stronger pressed-header profile");
}

void TestAppThemeSelectionResolutionIsSharedAndDeterministic()
{
    Require(ThemeModeFromThemeId(L"builtin/light") == ThemeMode::Light && ThemeModeFromThemeId(L"builtin/dark") == ThemeMode::Dark &&
                ThemeModeFromThemeId(L"builtin/rainbow") == ThemeMode::Rainbow && ThemeModeFromThemeId(L"builtin/highContrast") == ThemeMode::HighContrast &&
                ThemeModeFromThemeId(L"user/unknown") == ThemeMode::System,
            "theme-id mapping keeps the built-in identifiers and system fallback stable");

    Common::Settings::ThemeDefinition custom;
    custom.id          = L"user/dxui-selection-resolution";
    custom.name        = L"DxUi selection resolution";
    custom.baseThemeId = L"builtin/dark";
    custom.colors.emplace(L"app.accent", Common::Settings::ThemeColorSource(0x80402010u));
    custom.colors.emplace(L"window.background", Common::Settings::ThemeColorSource(0xFF112233u));
    custom.colors.emplace(L"navigation.background", Common::Settings::ThemeColorSource(0xFF203040u));
    custom.colors.emplace(L"navigation.separator", Common::Settings::ThemeColorSource(0xFF506070u));
    custom.colors.emplace(L"navigation.progressOk", Common::Settings::ThemeColorSource(0xFF109020u));

    const AppThemeSelectionResolution resolved = ResolveAppThemeSelection(custom.id, &custom, L"shared-theme-resolution-test");
    Require(resolved.customDefinitionResolved && resolved.resolvedColors.has_value() && resolved.baseMode == ThemeMode::Dark,
            "custom theme selection resolves once against its declared base mode");
    RequireColorNear(resolved.theme.accent,
                     D2D1::ColorF(0x40 / 255.0f, 0x20 / 255.0f, 0x10 / 255.0f, 0x80 / 255.0f),
                     "custom theme selection preserves ARGB accent channels and alpha");
    Require(resolved.theme.windowBackground == RGB(0x11, 0x22, 0x33), "custom theme selection applies the window background override");
    Require(resolved.theme.navigationView.gdiBackground == RGB(0x20, 0x30, 0x40) && resolved.theme.navigationView.gdiBorder == RGB(0x20, 0x30, 0x40) &&
                resolved.theme.navigationView.gdiBorderPen == RGB(0x50, 0x60, 0x70),
            "custom theme selection keeps navigation GDI colors synchronized with resolved D2D overrides");
    RequireColorNear(resolved.theme.fileOperations.progressTotal,
                     D2D1::ColorF(0x10 / 255.0f, 0x90 / 255.0f, 0x20 / 255.0f),
                     "custom theme selection derives file-operation colors from the effective navigation theme");

    Common::Settings::ThemeDefinition invalid = custom;
    invalid.id                                = L"user/dxui-selection-invalid";
    Common::Settings::ThemeColorSource missingReference;
    missingReference.kind         = Common::Settings::ThemeColorSourceKind::Reference;
    missingReference.references   = {L"missing.semantic.color"};
    invalid.colors[L"app.accent"] = std::move(missingReference);

    const AppThemeSelectionResolution rejected = ResolveAppThemeSelection(invalid.id, &invalid, L"shared-theme-resolution-test");
    Require(! rejected.customDefinitionResolved && ! rejected.resolvedColors.has_value() && rejected.baseMode == ThemeMode::Dark,
            "invalid custom theme selection reports failure without retaining partial resolved colors");
    const AppTheme darkBase = ResolveAppTheme(ThemeMode::Dark, L"shared-theme-resolution-test");
    RequireColorNear(rejected.theme.accent, darkBase.accent, "invalid custom theme selection falls back to the unmodified base theme");
}

void TestViewerThemeAccentPressedUsesBaseThemePolarity()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                 = 2u;
    viewerTheme.backgroundArgb          = 0xFFF2F2F2u;
    viewerTheme.textArgb                = 0xFF202020u;
    viewerTheme.selectionBackgroundArgb = 0xFFD8E8FFu;
    viewerTheme.selectionTextArgb       = 0xFF102030u;
    viewerTheme.accentArgb              = 0xFF808080u;
    viewerTheme.darkMode                = TRUE;
    viewerTheme.darkBase                = FALSE;

    const ThemePalette palette = MakeThemePaletteFromViewerTheme(viewerTheme);
    const auto luminance       = [](const D2D1_COLOR_F& color) noexcept { return (0.2126f * color.r) + (0.7152f * color.g) + (0.0722f * color.b); };

    Require(palette.dark, "viewer theme can request dark UI mode with a light base theme");
    Require(luminance(palette.accentPressed) < luminance(palette.accent),
            "light-base viewer theme derives a darker pressed accent even when UI dark mode is enabled");
}

void TestNewTokensPresentInLightDefaultPalette()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette light = MakeDefaultThemePalette(false);

    Require(light.smokeOverlay.a > 0.0f, "light default palette smoke overlay has non-zero alpha");
    RequireColorDifferent(light.borderDefault, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "light default palette border default is non-zero");
    RequireColorDifferent(light.borderStrong, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "light default palette border strong is non-zero");
    RequireColorDifferent(light.accentHover, light.accent, "light default palette accent hover differs from accent");
    RequireColorDifferent(light.accentPressed, light.accent, "light default palette accent pressed differs from accent");
    RequireColorDifferent(light.focusStrokeOuter, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "light default palette focus stroke outer is non-zero");
    RequireColorDifferent(light.focusStrokeInner, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "light default palette focus stroke inner is non-zero");
    RequireColorDifferent(light.focusStrokeOuter, light.focusStrokeInner, "light default palette focus stroke outer and inner differ");
}

void TestRadioButtonVisualStyleDerivesFromThemePalette()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme = MakeDefaultThemePalette(true);

    const RadioButtonVisualStyle idle    = ResolveRadioButtonVisualStyle(theme, true, false, false, false, false, false);
    const RadioButtonVisualStyle hovered = ResolveRadioButtonVisualStyle(theme, true, true, false, false, false, false);
    const RadioButtonVisualStyle checked = ResolveRadioButtonVisualStyle(theme, true, false, false, false, false, true);

    RequireColorNear(idle.text, theme.text, "radio button idle text uses the palette text token");
    Require(! idle.showHoverFill, "radio button idle state does not show hover fill");
    Require(hovered.showHoverFill, "radio button hovered state shows hover fill");
    RequireColorDifferent(checked.circleFill, idle.circleFill, "radio button checked state has different circle fill than idle");
    RequireFloatNear(checked.dotDiameterDip, 8.0f, 0.0001f, "radio button checked dot diameter defaults to 8 DIP");
}

void TestProgressBarVisualStyleUsesSelectionFill()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme           = MakeDefaultThemePalette(true);
    const ProgressBarVisualStyle style = ResolveProgressBarVisualStyle(theme);

    RequireColorNear(style.progressFill, theme.selectionFill, "progress bar visual style uses selection fill for progress");
    RequireColorDifferent(style.trackFill, style.progressFill, "progress bar track fill differs from progress fill");
}

void TestToolbarVisualStyleUsesCardBackground()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme       = MakeDefaultThemePalette(true);
    const ToolbarVisualStyle style = ResolveToolbarVisualStyle(theme);

    RequireColorNear(style.background, theme.cardBackground, "toolbar visual style uses card background token");
    RequireColorDifferent(style.bottomBorder, style.background, "toolbar bottom border differs from background");
    RequireColorDifferent(style.separatorLine, style.background, "toolbar separator differs from background");
}

} // namespace

void RunThemeTests()
{
    TestColorFromArgbPreservesAlpha();
    TestSharedColorRefArgbAndTruncatingBlendGoldenValues();
    TestStableVisualHash32Utf16V1GoldenValuesAndNamedEncodingPolicy();
    TestHsvColorRefNegativeHuePoliciesRemainExplicitAndDistinct();
    TestSharedLuminanceMathSeparatesLinearizedAndEncodedPolicies();
    TestD2dBlendClampsInterpolationAmount();
    TestWindowSizingDpiRoundingAndPopupGeometryGoldenValues();
    TestViewerFileComboPopupInputAndHeightGoldenValues();
    TestViewerTitleBarThemeGoldenValues();
    TestModifierCompositionCoversEveryCombination();
    TestUtfConversionPoliciesRemainExplicitAndPreserveEmbeddedNulls();
    TestOrdinalTrimAndEnvironmentPoliciesRemainExplicit();
    TestWindowsPathClassesKeepDriveRelativeAndFullyAbsolutePoliciesDistinct();
    TestBinaryThroughputGrammarPreservesBoundaryPoliciesAndRoundTrips();
    TestCompactFileMetadataFormattingPreservesTimeAndAttributePolicies();
    TestPercentEncodingAndHandleIoPoliciesCoverForwardProgressAndAtomicSiblings();
    TestCloudPaginationGuardBoundsProgressAndCancellation();
    TestYyjsonMemberPoliciesDistinguishPresenceTypeRangeAndCoercion();
    TestViewerThemePaletteDerivesDarkControlChrome();
    TestViewerThemePaletteDerivesLightControlChrome();
    TestViewerThemePaletteDerivesDarkHighContrastChrome();
    TestListIconColorUsesSelectionFillChrome();
    TestListIconColorUsesViewerDerivedSelectionFillChrome();
    TestGridBusyColorUsesSelectionFillChrome();
    TestGridBusyColorUsesViewerDerivedSelectionFillChrome();
    TestGridProgressVisualStyleUsesSharedTrackAndFillChrome();
    TestGridProgressVisualStyleUsesViewerDerivedTrackAndFillChrome();
    TestGridProgressVisualStyleUsesHighContrastTrackAndFillChrome();
    TestGridCheckboxVisualStyleUsesSharedCheckboxChrome();
    TestGridCheckboxVisualStyleUsesViewerDerivedSelectionTextChrome();
    TestGridSwatchVisualStyleUsesSharedRowChrome();
    TestGridSwatchVisualStyleUsesViewerDerivedRowChrome();
    TestGridSwatchVisualStyleUsesHighContrastRowChrome();
    TestGridBadgeVisualStyleUsesSharedRowChrome();
    TestGridBadgeVisualStyleUsesViewerDerivedRowChrome();
    TestGridBadgeVisualStyleUsesHighContrastRowChrome();
    TestTreeBadgeVisualStyleUsesSharedAdornmentChrome();
    TestTreeBadgeVisualStyleUsesViewerDerivedAdornmentChrome();
    TestTreeBadgeVisualStyleUsesHighContrastAdornmentChrome();
    TestButtonVisualStyleMatchesPreferencesFlatChrome();
    TestPrimaryButtonVisualStyleUsesAccentChrome();
    TestPrimaryButtonVisualStyleFallsBackToReadableText();
    TestButtonHighContrastFocusRingStaysVisibleWithoutKeyboardFocus();
    TestButtonHighContrastDisabledBorderStaysVisible();
    TestButtonVisualStyleInterpolatesHoverAndFocusStrength();
    TestPrimaryButtonVisualStyleUsesViewerDerivedSelectionTextChrome();
    TestButtonVisualStyleUsesViewerDerivedPressedBorderChrome();
    TestButtonVisualStyleUsesViewerDerivedHotAndPressedFillChrome();
    TestAdornmentColorsUsePaletteChrome();
    TestAdornmentColorsUseViewerDerivedSelectionAndAlertChrome();
    TestToggleVisualStyleMatchesPreferencesSwitchChrome();
    TestToggleVisualStyleUsesViewerDerivedKnobChrome();
    TestToggleHighContrastDisabledBordersStayVisible();
    TestToggleVisualStyleInterpolatesHoverStrength();
    TestToggleLayoutMetricsUseCompactSwitchOnlyChromeWhenTextIsEmpty();
    TestColorSwatchVisualStyleUsesSharedOverlayChrome();
    TestColorSwatchUsesViewerDerivedOverlayChrome();
    TestColorSwatchHighContrastFocusRingStaysVisibleWithoutKeyboardFocus();
    TestCardPanelVisualStyleUsesSharedSurfaceChrome();
    TestCardPanelUsesViewerDerivedSurfaceChrome();
    TestCardPanelHighContrastUsesPaletteSurfaceChrome();
    TestGridSurfaceVisualStyleUsesSharedGridChrome();
    TestGridSurfaceVisualStyleUsesViewerDerivedGridChrome();
    TestGridSurfaceHighContrastUsesPaletteGridChrome();
    TestTreeSurfaceVisualStyleUsesSharedTreeChrome();
    TestTreeSurfaceVisualStyleUsesViewerDerivedTreeChrome();
    TestTreeSurfaceHighContrastUsesPaletteTreeChrome();
    TestGridHeaderVisualStyleUsesSharedHeaderChrome();
    TestGridHeaderVisualStyleUsesViewerDerivedHeaderChrome();
    TestGridHeaderHighContrastUsesPaletteHeaderChrome();
    TestTooltipVisualStyleUsesSharedTooltipChrome();
    TestTooltipUsesViewerDerivedTooltipChrome();
    TestTooltipHighContrastBorderStrengthensLegibility();
    TestCheckboxVisualStyleMatchesStandardCheckboxChrome();
    TestLabelVisualStyleUsesPaletteTextChrome();
    TestLabelVisualStyleUsesViewerDerivedTextChrome();
    TestLabelVisualStyleUsesHighContrastTextChrome();
    TestCheckboxVisualStyleInterpolatesHoverStrength();
    TestCheckboxVisualStyleUsesViewerDerivedSelectionTextChrome();
    TestCheckboxHighContrastDisabledBordersStayVisible();
    TestComboBoxVariantsExposeDistinctChromeContracts();
    TestComboBoxHoverStyleStrengthensFieldChrome();
    TestComboBoxModernPopupUsesViewerDerivedPressedChrome();
    TestComboBoxUsesViewerDerivedInputBorderChrome();
    TestComboBoxHighContrastFocusAccentStaysVisibleWithoutKeyboardFocus();
    TestComboBoxHighContrastDisabledBordersStayVisible();
    TestGridRowIconChromeFollowsResolvedRowVisuals();
    TestGridRowIconChromeFollowsViewerDerivedRowVisuals();
    TestGridCellChromeFollowsResolvedCellVisuals();
    TestGridCellChromeFollowsViewerDerivedResolvedCellVisuals();

    // Phase 0-4 new tokens and visual styles
    TestNewTokensPresentInDarkDefaultPalette();
    TestDarkDefaultPaletteOverridesLightInteractionTokens();
    TestAppThemeDxPaletteRefreshesAccentVariantsAfterAssigningAccent();
    TestFolderContentDxPalettePreservesItsNamedProfile();
    TestAppThemeSelectionResolutionIsSharedAndDeterministic();
    TestViewerThemeAccentPressedUsesBaseThemePolarity();
    TestNewTokensPresentInLightDefaultPalette();
    TestRadioButtonVisualStyleDerivesFromThemePalette();
    TestProgressBarVisualStyleUsesSelectionFill();
    TestToolbarVisualStyleUsesCardBackground();
}

#include "ViewerVLC.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstdint>
#include <cwctype>
#include <format>
#include <limits>
#include <optional>
#include <string_view>
#include <vector>

#include <commdlg.h>
#include <d2d1.h>
#include <d2d1helper.h>
#include <dwmapi.h>
#include <dwrite.h>
#include <shellapi.h>
#include <winreg.h>

#include <yyjson.h>

#include "Helpers.h"

#include "resource.h"

extern HINSTANCE g_hInstance;

#pragma comment(lib, "Dwmapi.lib")
#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "dwrite.lib")

// libVLC forward declarations (loaded dynamically; no headers required).
struct libvlc_instance_t;
struct libvlc_media_t;
struct libvlc_media_player_t;
using libvlc_time_t = int64_t;

namespace
{
constexpr UINT_PTR kUiTimerId          = 1;
constexpr UINT kUiTimerIntervalMs      = 200;
constexpr UINT_PTR kHudAnimTimerId     = 2;
constexpr UINT kHudAnimIntervalMs      = 16;
constexpr float kHudDimOpacity         = 0.18f;
constexpr ULONGLONG kHudIdleDimDelayMs = 20'000;

constexpr char kViewerVlcSchemaJson[] = R"json(
{
  "version": 1,
  "title": "VLC Viewer",
  "fields": [
    {
      "key": "vlcInstallPath",
      "label": "VLC installation folder",
      "type": "text",
      "default": "",
      "browse": "folder",
      "description": "Folder containing vlc.exe and libvlc.dll (typically: C:\\\\Program Files\\\\VideoLAN\\\\VLC)."
    },
    {
      "key": "autoDetectVlc",
      "label": "Auto-detect VLC",
      "type": "bool",
      "default": true,
      "description": "If enabled, the viewer will try common install locations when the path is empty."
    },
    {
      "key": "quiet",
      "label": "Quiet mode",
      "type": "bool",
      "default": true,
      "description": "Reduce VLC logging."
    },
    {
      "key": "defaultPlaybackRatePercent",
      "label": "Default playback speed (%)",
      "type": "value",
      "default": 100,
      "min": 25,
      "max": 400,
      "description": "Applied when a file is opened. 100 = normal speed."
    },
    {
      "key": "fileCachingMs",
      "label": "File caching (ms)",
      "type": "value",
      "default": 300,
      "min": 0,
      "max": 60000,
      "description": "Local file buffer. Increase if playback stutters on slow media."
    },
    {
      "key": "networkCachingMs",
      "label": "Network caching (ms)",
      "type": "value",
      "default": 1000,
      "min": 0,
      "max": 60000,
      "description": "Network stream buffer. Increase for unstable connections."
    },
    {
      "key": "avcodecHw",
      "label": "Hardware decoding",
      "type": "option",
      "default": "any",
      "description": "Decoder acceleration (maps to --avcodec-hw).",
      "options": [
        { "value": "any", "label": "Auto" },
        { "value": "none", "label": "Off" },
        { "value": "dxva2", "label": "DXVA2" },
        { "value": "d3d11va", "label": "D3D11VA" }
      ]
    },
    {
      "key": "audioVisualization",
      "label": "Audio visualization",
      "type": "option",
      "default": "visual",
      "description": "When opening audio-only files, show a visualizer (maps to --audio-visual).",
      "options": [
        { "value": "off", "label": "Off" },
        { "value": "any", "label": "Any visualization" },
        { "value": "projectm", "label": "ProjectM" },
        { "value": "spectrometer", "label": "Spectrometer" },
        { "value": "spectrum", "label": "Spectrum" },
        { "value": "vumeter", "label": "VU Meter" },
        { "value": "goom", "label": "Goom" },
        { "value": "glspectrum", "label": "3D Spectrum" },
        { "value": "visual", "label": "Visual" }
      ]
    },
    {
      "key": "videoOutput",
      "label": "Video output (vout)",
      "type": "text",
      "default": "",
      "description": "Optional override for --vout (example: direct3d11)."
    },
    {
      "key": "audioOutput",
      "label": "Audio output (aout)",
      "type": "text",
      "default": "",
      "description": "Optional override for --aout (example: mmdevice)."
    },
    {
      "key": "extraArgs",
      "label": "Extra VLC arguments",
      "type": "text",
      "default": "",
      "description": "Additional libVLC options (space-separated). Example: --no-sub-autodetect-file"
    }
  ]
}
)json";

template <typename T> [[nodiscard]] bool TryLoadProc(HMODULE module, const char* name, T& out) noexcept
{
#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    out = reinterpret_cast<T>(GetProcAddress(module, name));
#pragma warning(pop)
    return out != nullptr;
}

[[nodiscard]] COLORREF ColorRefFromArgb(uint32_t argb) noexcept
{
    const BYTE r = static_cast<BYTE>((argb >> 16) & 0xFFu);
    const BYTE g = static_cast<BYTE>((argb >> 8) & 0xFFu);
    const BYTE b = static_cast<BYTE>(argb & 0xFFu);
    return RGB(r, g, b);
}

[[nodiscard]] COLORREF BlendColor(COLORREF under, COLORREF over, uint8_t alpha) noexcept
{
    const uint32_t inv = static_cast<uint32_t>(255u - alpha);

    const uint32_t ur = static_cast<uint32_t>(GetRValue(under));
    const uint32_t ug = static_cast<uint32_t>(GetGValue(under));
    const uint32_t ub = static_cast<uint32_t>(GetBValue(under));

    const uint32_t or_ = static_cast<uint32_t>(GetRValue(over));
    const uint32_t og  = static_cast<uint32_t>(GetGValue(over));
    const uint32_t ob  = static_cast<uint32_t>(GetBValue(over));

    const uint8_t r = static_cast<uint8_t>((ur * inv + or_ * static_cast<uint32_t>(alpha)) / 255u);
    const uint8_t g = static_cast<uint8_t>((ug * inv + og * static_cast<uint32_t>(alpha)) / 255u);
    const uint8_t b = static_cast<uint8_t>((ub * inv + ob * static_cast<uint32_t>(alpha)) / 255u);
    return RGB(r, g, b);
}

[[nodiscard]] COLORREF ContrastingTextColor(COLORREF background) noexcept
{
    const uint32_t r    = static_cast<uint32_t>(GetRValue(background));
    const uint32_t g    = static_cast<uint32_t>(GetGValue(background));
    const uint32_t b    = static_cast<uint32_t>(GetBValue(background));
    const uint32_t luma = (r * 299u + g * 587u + b * 114u) / 1000u;
    return luma < 128u ? RGB(255, 255, 255) : RGB(0, 0, 0);
}

[[nodiscard]] uint32_t StableHash32(std::wstring_view text) noexcept
{
    uint32_t hash = 2166136261u;
    for (wchar_t ch : text)
    {
        hash ^= static_cast<uint32_t>(ch);
        hash *= 16777619u;
    }
    return hash;
}

[[nodiscard]] COLORREF ColorFromHSV(float hueDegrees, float saturation, float value) noexcept
{
    const float h = std::fmod(std::max(0.0f, hueDegrees), 360.0f);
    const float s = std::clamp(saturation, 0.0f, 1.0f);
    const float v = std::clamp(value, 0.0f, 1.0f);

    const float c = v * s;
    const float x = c * (1.0f - std::fabs(std::fmod(h / 60.0f, 2.0f) - 1.0f));
    const float m = v - c;

    float rf = 0.0f;
    float gf = 0.0f;
    float bf = 0.0f;

    if (h < 60.0f)
    {
        rf = c;
        gf = x;
        bf = 0.0f;
    }
    else if (h < 120.0f)
    {
        rf = x;
        gf = c;
        bf = 0.0f;
    }
    else if (h < 180.0f)
    {
        rf = 0.0f;
        gf = c;
        bf = x;
    }
    else if (h < 240.0f)
    {
        rf = 0.0f;
        gf = x;
        bf = c;
    }
    else if (h < 300.0f)
    {
        rf = x;
        gf = 0.0f;
        bf = c;
    }
    else
    {
        rf = c;
        gf = 0.0f;
        bf = x;
    }

    const auto toByte = [](float v01) noexcept
    {
        const float scaled = std::clamp(v01 * 255.0f, 0.0f, 255.0f);
        return static_cast<BYTE>(std::lround(scaled));
    };

    const BYTE r = toByte(rf + m);
    const BYTE g = toByte(gf + m);
    const BYTE b = toByte(bf + m);
    return RGB(r, g, b);
}

[[nodiscard]] COLORREF ResolveAccentColor(const ViewerTheme& theme, std::wstring_view seed) noexcept
{
    if (theme.rainbowMode)
    {
        const uint32_t h = StableHash32(seed);
        const float hue  = static_cast<float>(h % 360u);
        const float sat  = theme.darkBase ? 0.70f : 0.55f;
        const float val  = theme.darkBase ? 0.95f : 0.85f;
        return ColorFromHSV(hue, sat, val);
    }

    return ColorRefFromArgb(theme.accentArgb);
}

[[nodiscard]] D2D1_COLOR_F ColorFFromColorRef(COLORREF color, float alpha = 1.0f) noexcept
{
    const float r = static_cast<float>(GetRValue(color)) / 255.0f;
    const float g = static_cast<float>(GetGValue(color)) / 255.0f;
    const float b = static_cast<float>(GetBValue(color)) / 255.0f;
    return D2D1::ColorF(r, g, b, alpha);
}

[[nodiscard]] D2D1_RECT_F RectFFromRect(const RECT& rc) noexcept
{
    return D2D1::RectF(static_cast<float>(rc.left), static_cast<float>(rc.top), static_cast<float>(rc.right), static_cast<float>(rc.bottom));
}

struct HudLayout
{
    RECT play{};
    RECT stop{};
    RECT snapshot{};
    RECT seekHit{};
    RECT seekTrack{};
    RECT time{};
    RECT speed{};
    RECT volume{};
};

[[nodiscard]] HudLayout ComputeHudLayout(int width, int height, UINT dpi) noexcept
{
    HudLayout layout{};

    const auto px = [&](int dip) noexcept { return MulDiv(dip, static_cast<int>(dpi), 96); };

    const int inset  = px(12);
    const int gap    = px(10);
    const int btn    = px(36);
    const int timeW  = px(140);
    const int trackH = std::max(1, px(6));

    const int y = std::max(0, (height - btn) / 2);

    int x       = inset;
    layout.play = {x, y, x + btn, y + btn};
    x += btn + gap;
    layout.stop = {x, y, x + btn, y + btn};
    x += btn + gap;
    layout.snapshot = {x, y, x + btn, y + btn};
    x += btn + gap;

    int right     = std::max(x, width - inset);
    layout.volume = {std::max(x, right - btn), y, right, y + btn};
    right         = std::max<int>(x, static_cast<int>(layout.volume.left) - gap);

    layout.speed = {std::max(x, right - btn), y, right, y + btn};
    right        = std::max<int>(x, static_cast<int>(layout.speed.left) - gap);

    const bool showTime = (right - x) >= (timeW + px(80));
    if (showTime)
    {
        layout.time = {std::max(x, right - timeW), y, right, y + btn};
        right       = std::max<int>(x, static_cast<int>(layout.time.left) - gap);
    }
    else
    {
        layout.time = {};
    }

    layout.seekHit   = {x, y, std::max(x, right), y + btn};
    const int trackY = y + (btn - trackH) / 2;
    layout.seekTrack = {x, trackY, std::max(x, right), trackY + trackH};

    return layout;
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty() || text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return {};
    }

    const int required = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty() || text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return {};
    }

    const int required = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
    {
        return {};
    }

    std::string result(static_cast<size_t>(required), '\0');
    const int written =
        WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required, nullptr, nullptr);
    if (written != required)
    {
        return {};
    }

    return result;
}

[[nodiscard]] bool IsRegularFile(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    std::error_code ec;
    return std::filesystem::is_regular_file(path, ec);
}

[[nodiscard]] bool IsDirectory(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    std::error_code ec;
    return std::filesystem::is_directory(path, ec);
}

[[nodiscard]] std::filesystem::path NormalizeVlcInstallPath(const std::filesystem::path& input) noexcept
{
    if (input.empty())
    {
        return {};
    }

    if (IsRegularFile(input))
    {
        const std::wstring ext = input.extension().wstring();
        if (_wcsicmp(ext.c_str(), L".exe") == 0 || _wcsicmp(ext.c_str(), L".dll") == 0)
        {
            return input.parent_path();
        }
    }

    return input;
}

[[nodiscard]] bool IsVlcInstallDir(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    return IsRegularFile(path / L"libvlc.dll") && IsRegularFile(path / L"vlc.exe") && IsDirectory(path / L"plugins");
}

[[nodiscard]] std::optional<std::filesystem::path> TryReadRegPath(HKEY root, std::wstring_view subKey, std::wstring_view valueName) noexcept
{
    const std::wstring subKeyStr(subKey);
    const std::wstring valueStr(valueName);
    const wchar_t* valuePtr = valueStr.empty() ? nullptr : valueStr.c_str();

    DWORD type          = 0;
    DWORD bytes         = 0;
    const LSTATUS first = RegGetValueW(root, subKeyStr.c_str(), valuePtr, RRF_RT_REG_SZ | RRF_RT_REG_EXPAND_SZ, &type, nullptr, &bytes);
    if (first != ERROR_SUCCESS || bytes < sizeof(wchar_t))
    {
        return std::nullopt;
    }

    std::wstring buffer(static_cast<size_t>(bytes / sizeof(wchar_t)), L'\0');
    const LSTATUS second = RegGetValueW(root, subKeyStr.c_str(), valuePtr, RRF_RT_REG_SZ | RRF_RT_REG_EXPAND_SZ, &type, buffer.data(), &bytes);
    if (second != ERROR_SUCCESS || bytes < sizeof(wchar_t))
    {
        return std::nullopt;
    }

    const size_t wcharCount = static_cast<size_t>(bytes / sizeof(wchar_t));
    buffer.resize(wcharCount);

    if (! buffer.empty() && buffer.back() == L'\0')
    {
        buffer.pop_back();
    }

    if (buffer.empty())
    {
        return std::nullopt;
    }

    if (type == REG_EXPAND_SZ)
    {
        const DWORD needed = ExpandEnvironmentStringsW(buffer.c_str(), nullptr, 0);
        if (needed > 0 && needed < 32768)
        {
            std::wstring expanded(static_cast<size_t>(needed), L'\0');
            const DWORD written = ExpandEnvironmentStringsW(buffer.c_str(), expanded.data(), needed);
            if (written > 0 && written <= needed)
            {
                if (! expanded.empty() && expanded.back() == L'\0')
                {
                    expanded.pop_back();
                }
                if (! expanded.empty())
                {
                    buffer = std::move(expanded);
                }
            }
        }
    }

    return std::filesystem::path(buffer);
}

[[nodiscard]] std::optional<std::filesystem::path> TryGetEnvPath(const wchar_t* var) noexcept
{
    if (! var || var[0] == L'\0')
    {
        return std::nullopt;
    }

    const DWORD required = GetEnvironmentVariableW(var, nullptr, 0);
    if (required == 0 || required > 32768)
    {
        return std::nullopt;
    }

    std::wstring value(static_cast<size_t>(required), L'\0');
    const DWORD written = GetEnvironmentVariableW(var, value.data(), required);
    if (written == 0 || written >= required)
    {
        return std::nullopt;
    }

    value.resize(static_cast<size_t>(written));
    if (value.empty())
    {
        return std::nullopt;
    }

    return std::filesystem::path(value);
}

[[nodiscard]] std::optional<std::filesystem::path> AutoDetectVlcInstallDir() noexcept
{
    for (HKEY root : {HKEY_LOCAL_MACHINE, HKEY_CURRENT_USER})
    {
        if (auto p = TryReadRegPath(root, L"SOFTWARE\\VideoLAN\\VLC", L"InstallDir"); p.has_value())
        {
            const std::filesystem::path dir = NormalizeVlcInstallPath(p.value());
            if (IsVlcInstallDir(dir))
            {
                return dir;
            }
        }

        if (auto p = TryReadRegPath(root, L"SOFTWARE\\Wow6432Node\\VideoLAN\\VLC", L"InstallDir"); p.has_value())
        {
            const std::filesystem::path dir = NormalizeVlcInstallPath(p.value());
            if (IsVlcInstallDir(dir))
            {
                return dir;
            }
        }

        if (auto p = TryReadRegPath(root, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\vlc.exe", L""); p.has_value())
        {
            const std::filesystem::path dir = NormalizeVlcInstallPath(p.value());
            if (IsVlcInstallDir(dir))
            {
                return dir;
            }
        }
    }

    if (const auto programFiles = TryGetEnvPath(L"ProgramFiles"); programFiles.has_value())
    {
        const std::filesystem::path dir = programFiles.value() / L"VideoLAN" / L"VLC";
        if (IsVlcInstallDir(dir))
        {
            return dir;
        }
    }

    if (const auto programFilesX86 = TryGetEnvPath(L"ProgramFiles(x86)"); programFilesX86.has_value())
    {
        const std::filesystem::path dir = programFilesX86.value() / L"VideoLAN" / L"VLC";
        if (IsVlcInstallDir(dir))
        {
            return dir;
        }
    }

    const DWORD probe = SearchPathW(nullptr, L"vlc.exe", nullptr, 0, nullptr, nullptr);
    if (probe > 0 && probe < 32768)
    {
        std::wstring buffer(static_cast<size_t>(probe + 1), L'\0');
        const DWORD written = SearchPathW(nullptr, L"vlc.exe", nullptr, probe + 1, buffer.data(), nullptr);
        if (written > 0)
        {
            buffer.resize(static_cast<size_t>(written));
            const std::filesystem::path dir = NormalizeVlcInstallPath(std::filesystem::path(buffer));
            if (IsVlcInstallDir(dir))
            {
                return dir;
            }
        }
    }

    return std::nullopt;
}

[[nodiscard]] bool EqualsIgnoreCase(std::wstring_view a, std::wstring_view b) noexcept
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

[[nodiscard]] bool IsAudioExtension(std::wstring_view ext) noexcept
{
    static constexpr std::array<std::wstring_view, 11> kAudioExts{{
        L".m4a",
        L".mp3",
        L".aac",
        L".flac",
        L".wav",
        L".ogg",
        L".opus",
        L".wma",
        L".mka",
        L".aif",
        L".aiff",
    }};

    for (const auto& e : kAudioExts)
    {
        if (EqualsIgnoreCase(ext, e))
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] std::vector<std::string> SplitVlcArgs(std::string_view text) noexcept
{
    std::vector<std::string> args;
    std::string current;
    bool inQuotes  = false;
    char quoteChar = '\0';
    bool escaping  = false;

    const auto flush = [&]
    {
        if (! current.empty())
        {
            args.push_back(std::move(current));
            current.clear();
        }
    };

    for (const char ch : text)
    {
        const bool isSpace = (ch == ' ') || (ch == '\t') || (ch == '\r') || (ch == '\n');
        if (! inQuotes && isSpace)
        {
            flush();
            continue;
        }

        if (escaping)
        {
            current.push_back(ch);
            escaping = false;
            continue;
        }

        if (inQuotes)
        {
            if (ch == quoteChar)
            {
                inQuotes = false;
                continue;
            }

            if (ch == '\\')
            {
                escaping = true;
                continue;
            }

            current.push_back(ch);
            continue;
        }

        if (ch == '"' || ch == '\'')
        {
            inQuotes  = true;
            quoteChar = ch;
            continue;
        }

        current.push_back(ch);
    }

    if (escaping)
    {
        current.push_back('\\');
    }

    flush();
    return args;
}

[[nodiscard]] std::wstring FormatDurationMs(libvlc_time_t ms) noexcept
{
    if (ms <= 0)
    {
        return L"--:--";
    }

    const int64_t totalSeconds = ms / 1000;
    const int64_t seconds      = totalSeconds % 60;
    const int64_t minutes      = (totalSeconds / 60) % 60;
    const int64_t hours        = totalSeconds / 3600;

    if (hours > 0)
    {
        return std::format(L"{:d}:{:02d}:{:02d}", hours, minutes, seconds);
    }

    return std::format(L"{:02d}:{:02d}", minutes, seconds);
}

[[nodiscard]] std::wstring FormatPlaybackRate(float rate) noexcept
{
    const float clamped = std::clamp(rate, 0.25f, 4.0f);
    const float r0      = std::round(clamped);
    if (std::fabs(clamped - r0) < 0.001f)
    {
        return std::format(L"{:.0f}×", clamped);
    }

    const float r1 = std::round(clamped * 2.0f) / 2.0f;
    if (std::fabs(clamped - r1) < 0.001f)
    {
        return std::format(L"{:.1f}×", clamped);
    }

    return std::format(L"{:.2f}×", clamped);
}
} // namespace

struct VlcState
{
    VlcState()                           = default;
    VlcState(const VlcState&)            = delete;
    VlcState(VlcState&&)                 = delete;
    VlcState& operator=(const VlcState&) = delete;
    VlcState& operator=(VlcState&&)      = delete;

    ~VlcState()
    {
        if (! dllDirectoryWasSet)
        {
            return;
        }

        if (previousDllDirectory.empty())
        {
            SetDllDirectoryW(nullptr);
            return;
        }

        SetDllDirectoryW(previousDllDirectory.c_str());
    }

    wil::unique_hmodule module;

    libvlc_instance_t*(__cdecl* libvlc_new)(int, const char* const*)                                            = nullptr;
    void(__cdecl* libvlc_release)(libvlc_instance_t*)                                                           = nullptr;
    libvlc_media_t*(__cdecl* libvlc_media_new_path)(libvlc_instance_t*, const char*)                            = nullptr;
    void(__cdecl* libvlc_media_release)(libvlc_media_t*)                                                        = nullptr;
    libvlc_media_player_t*(__cdecl* libvlc_media_player_new_from_media)(libvlc_media_t*)                        = nullptr;
    void(__cdecl* libvlc_media_player_release)(libvlc_media_player_t*)                                          = nullptr;
    void(__cdecl* libvlc_media_player_set_hwnd)(libvlc_media_player_t*, void*)                                  = nullptr;
    int(__cdecl* libvlc_media_player_play)(libvlc_media_player_t*)                                              = nullptr;
    void(__cdecl* libvlc_media_player_pause)(libvlc_media_player_t*)                                            = nullptr;
    void(__cdecl* libvlc_media_player_stop)(libvlc_media_player_t*)                                             = nullptr;
    int(__cdecl* libvlc_media_player_is_playing)(libvlc_media_player_t*)                                        = nullptr;
    libvlc_time_t(__cdecl* libvlc_media_player_get_time)(libvlc_media_player_t*)                                = nullptr;
    void(__cdecl* libvlc_media_player_set_time)(libvlc_media_player_t*, libvlc_time_t)                          = nullptr;
    libvlc_time_t(__cdecl* libvlc_media_player_get_length)(libvlc_media_player_t*)                              = nullptr;
    int(__cdecl* libvlc_media_player_set_rate)(libvlc_media_player_t*, float)                                   = nullptr;
    float(__cdecl* libvlc_media_player_get_rate)(libvlc_media_player_t*)                                        = nullptr;
    int(__cdecl* libvlc_audio_set_volume)(libvlc_media_player_t*, int)                                          = nullptr;
    int(__cdecl* libvlc_audio_get_volume)(libvlc_media_player_t*)                                               = nullptr;
    int(__cdecl* libvlc_video_take_snapshot)(libvlc_media_player_t*, unsigned, const char*, unsigned, unsigned) = nullptr;

    struct InstanceDeleter
    {
        void(__cdecl* release)(libvlc_instance_t*) = nullptr;
        void operator()(libvlc_instance_t* p) const noexcept
        {
            if (p && release)
            {
                release(p);
            }
        }
    };

    struct MediaDeleter
    {
        void(__cdecl* release)(libvlc_media_t*) = nullptr;
        void operator()(libvlc_media_t* p) const noexcept
        {
            if (p && release)
            {
                release(p);
            }
        }
    };

    struct PlayerDeleter
    {
        void(__cdecl* release)(libvlc_media_player_t*) = nullptr;
        void operator()(libvlc_media_player_t* p) const noexcept
        {
            if (p && release)
            {
                release(p);
            }
        }
    };

    std::unique_ptr<libvlc_instance_t, InstanceDeleter> instance{nullptr, {}};
    std::unique_ptr<libvlc_media_player_t, PlayerDeleter> player{nullptr, {}};

    std::wstring previousDllDirectory;
    bool dllDirectoryWasSet = false;

    std::filesystem::path installDir;
    std::string instanceArgsKey;
};

ViewerVLC::ViewerVLC()
{
    _metaId          = L"builtin/viewer-vlc";
    _metaShortId     = L"viewvlc";
    _metaName        = LoadStringResource(g_hInstance, IDS_VIEWERVLC_NAME);
    _metaDescription = LoadStringResource(g_hInstance, IDS_VIEWERVLC_DESCRIPTION);

    _metaData.id          = _metaId.c_str();
    _metaData.shortId     = _metaShortId.c_str();
    _metaData.name        = _metaName.empty() ? nullptr : _metaName.c_str();
    _metaData.description = _metaDescription.empty() ? nullptr : _metaDescription.c_str();
    _metaData.author      = nullptr;
    _metaData.version     = nullptr;

    static_cast<void>(SetConfiguration(nullptr));
}

ViewerVLC::~ViewerVLC() = default;

void ViewerVLC::SetHost(IHost* host) noexcept
{
    _hostAlerts = nullptr;

    if (! host)
    {
        return;
    }

    wil::com_ptr<IHostAlerts> alerts;
    const HRESULT hr = host->QueryInterface(__uuidof(IHostAlerts), alerts.put_void());
    if (SUCCEEDED(hr) && alerts)
    {
        _hostAlerts = std::move(alerts);
    }
}

HRESULT STDMETHODCALLTYPE ViewerVLC::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (ppvObject == nullptr)
    {
        return E_POINTER;
    }

    *ppvObject = nullptr;

    if (riid == __uuidof(IUnknown) || riid == __uuidof(IViewer))
    {
        *ppvObject = static_cast<IViewer*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IInformations))
    {
        *ppvObject = static_cast<IInformations*>(this);
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE ViewerVLC::AddRef() noexcept
{
    return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
}

ULONG STDMETHODCALLTYPE ViewerVLC::Release() noexcept
{
    const ULONG refs = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
    if (refs == 0)
    {
        delete this;
    }
    return refs;
}

HRESULT STDMETHODCALLTYPE ViewerVLC::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (metaData == nullptr)
    {
        return E_POINTER;
    }

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerVLC::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = kViewerVlcSchemaJson;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerVLC::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    std::wstring vlcInstallPath;
    bool autoDetectVlc                  = true;
    bool quiet                          = true;
    uint32_t fileCachingMs              = 300;
    uint32_t networkCachingMs           = 1000;
    uint32_t defaultPlaybackRatePercent = 100;
    std::string avcodecHw               = "any";
    std::string videoOutput;
    std::string audioOutput;
    std::string audioVisualization = "visual";
    std::string extraArgs;

    if (configurationJsonUtf8 != nullptr && configurationJsonUtf8[0] != '\0')
    {
        yyjson_doc* doc = yyjson_read(configurationJsonUtf8, strlen(configurationJsonUtf8), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
        if (doc)
        {
            auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

            yyjson_val* root = yyjson_doc_get_root(doc);
            if (root && yyjson_is_obj(root))
            {
                if (yyjson_val* v = yyjson_obj_get(root, "vlcInstallPath"); v && yyjson_is_str(v))
                {
                    const char* s = yyjson_get_str(v);
                    if (s)
                    {
                        vlcInstallPath = Utf16FromUtf8(s);
                    }
                }

                if (yyjson_val* v = yyjson_obj_get(root, "autoDetectVlc"); v)
                {
                    if (yyjson_is_bool(v))
                    {
                        autoDetectVlc = yyjson_get_bool(v);
                    }
                    else if (yyjson_is_sint(v))
                    {
                        autoDetectVlc = yyjson_get_sint(v) != 0;
                    }
                    else if (yyjson_is_uint(v))
                    {
                        autoDetectVlc = yyjson_get_uint(v) != 0;
                    }
                }

                if (yyjson_val* v = yyjson_obj_get(root, "quiet"); v)
                {
                    if (yyjson_is_bool(v))
                    {
                        quiet = yyjson_get_bool(v);
                    }
                    else if (yyjson_is_sint(v))
                    {
                        quiet = yyjson_get_sint(v) != 0;
                    }
                    else if (yyjson_is_uint(v))
                    {
                        quiet = yyjson_get_uint(v) != 0;
                    }
                }

                if (yyjson_val* v = yyjson_obj_get(root, "fileCachingMs"); v && (yyjson_is_sint(v) || yyjson_is_uint(v)))
                {
                    const int64_t raw = yyjson_is_sint(v) ? yyjson_get_sint(v) : static_cast<int64_t>(yyjson_get_uint(v));
                    fileCachingMs     = static_cast<uint32_t>(std::clamp<int64_t>(raw, 0, 60'000));
                }

                if (yyjson_val* v = yyjson_obj_get(root, "networkCachingMs"); v && (yyjson_is_sint(v) || yyjson_is_uint(v)))
                {
                    const int64_t raw = yyjson_is_sint(v) ? yyjson_get_sint(v) : static_cast<int64_t>(yyjson_get_uint(v));
                    networkCachingMs  = static_cast<uint32_t>(std::clamp<int64_t>(raw, 0, 60'000));
                }

                if (yyjson_val* v = yyjson_obj_get(root, "defaultPlaybackRatePercent"); v && (yyjson_is_sint(v) || yyjson_is_uint(v)))
                {
                    const int64_t raw          = yyjson_is_sint(v) ? yyjson_get_sint(v) : static_cast<int64_t>(yyjson_get_uint(v));
                    defaultPlaybackRatePercent = static_cast<uint32_t>(std::clamp<int64_t>(raw, 25, 400));
                }

                if (yyjson_val* v = yyjson_obj_get(root, "avcodecHw"); v && yyjson_is_str(v))
                {
                    if (const char* s = yyjson_get_str(v); s)
                    {
                        avcodecHw = s;
                    }
                }

                if (yyjson_val* v = yyjson_obj_get(root, "videoOutput"); v && yyjson_is_str(v))
                {
                    if (const char* s = yyjson_get_str(v); s)
                    {
                        videoOutput = s;
                    }
                }

                if (yyjson_val* v = yyjson_obj_get(root, "audioOutput"); v && yyjson_is_str(v))
                {
                    if (const char* s = yyjson_get_str(v); s)
                    {
                        audioOutput = s;
                    }
                }

                if (yyjson_val* v = yyjson_obj_get(root, "audioVisualization"); v && yyjson_is_str(v))
                {
                    if (const char* s = yyjson_get_str(v); s)
                    {
                        audioVisualization = s;
                    }
                }

                if (yyjson_val* v = yyjson_obj_get(root, "extraArgs"); v && yyjson_is_str(v))
                {
                    if (const char* s = yyjson_get_str(v); s)
                    {
                        extraArgs = s;
                    }
                }
            }
        }
    }

    _config.vlcInstallPath             = std::filesystem::path(vlcInstallPath);
    _config.autoDetectVlc              = autoDetectVlc;
    _config.quiet                      = quiet;
    _config.fileCachingMs              = fileCachingMs;
    _config.networkCachingMs           = networkCachingMs;
    _config.defaultPlaybackRatePercent = defaultPlaybackRatePercent;
    _config.avcodecHw                  = avcodecHw;
    _config.videoOutput                = videoOutput;
    _config.audioOutput                = audioOutput;
    _config.audioVisualization         = audioVisualization;
    _config.extraArgs                  = extraArgs;

    _hudRate = std::clamp(static_cast<float>(_config.defaultPlaybackRatePercent) / 100.0f, 0.25f, 4.0f);

    yyjson_mut_doc* doc = yyjson_mut_doc_new(nullptr);
    if (! doc)
    {
        return E_OUTOFMEMORY;
    }

    auto freeDoc = wil::scope_exit([&] { yyjson_mut_doc_free(doc); });

    yyjson_mut_val* root = yyjson_mut_obj(doc);
    if (! root)
    {
        return E_OUTOFMEMORY;
    }
    yyjson_mut_doc_set_root(doc, root);

    const std::string keyPathUtf8   = "vlcInstallPath";
    const std::string valuePathUtf8 = Utf8FromUtf16(vlcInstallPath);
    yyjson_mut_val* pathKey         = yyjson_mut_strncpy(doc, keyPathUtf8.c_str(), keyPathUtf8.size());
    yyjson_mut_val* pathVal         = yyjson_mut_strncpy(doc, valuePathUtf8.c_str(), valuePathUtf8.size());
    if (! pathKey || ! pathVal)
    {
        return E_OUTOFMEMORY;
    }
    static_cast<void>(yyjson_mut_obj_add(root, pathKey, pathVal));

    yyjson_mut_val* autoKey = yyjson_mut_strncpy(doc, "autoDetectVlc", strlen("autoDetectVlc"));
    yyjson_mut_val* autoVal = yyjson_mut_bool(doc, autoDetectVlc);
    if (! autoKey || ! autoVal)
    {
        return E_OUTOFMEMORY;
    }
    static_cast<void>(yyjson_mut_obj_add(root, autoKey, autoVal));

    yyjson_mut_val* quietKey = yyjson_mut_strncpy(doc, "quiet", strlen("quiet"));
    yyjson_mut_val* quietVal = yyjson_mut_bool(doc, quiet);
    if (! quietKey || ! quietVal)
    {
        return E_OUTOFMEMORY;
    }
    static_cast<void>(yyjson_mut_obj_add(root, quietKey, quietVal));

    yyjson_mut_val* rateKey = yyjson_mut_strncpy(doc, "defaultPlaybackRatePercent", strlen("defaultPlaybackRatePercent"));
    yyjson_mut_val* rateVal = yyjson_mut_uint(doc, defaultPlaybackRatePercent);
    if (! rateKey || ! rateVal)
    {
        return E_OUTOFMEMORY;
    }
    static_cast<void>(yyjson_mut_obj_add(root, rateKey, rateVal));

    yyjson_mut_val* fileCacheKey = yyjson_mut_strncpy(doc, "fileCachingMs", strlen("fileCachingMs"));
    yyjson_mut_val* fileCacheVal = yyjson_mut_uint(doc, fileCachingMs);
    if (! fileCacheKey || ! fileCacheVal)
    {
        return E_OUTOFMEMORY;
    }
    static_cast<void>(yyjson_mut_obj_add(root, fileCacheKey, fileCacheVal));

    yyjson_mut_val* netCacheKey = yyjson_mut_strncpy(doc, "networkCachingMs", strlen("networkCachingMs"));
    yyjson_mut_val* netCacheVal = yyjson_mut_uint(doc, networkCachingMs);
    if (! netCacheKey || ! netCacheVal)
    {
        return E_OUTOFMEMORY;
    }
    static_cast<void>(yyjson_mut_obj_add(root, netCacheKey, netCacheVal));

    yyjson_mut_val* hwKey = yyjson_mut_strncpy(doc, "avcodecHw", strlen("avcodecHw"));
    yyjson_mut_val* hwVal = yyjson_mut_strncpy(doc, avcodecHw.c_str(), avcodecHw.size());
    if (! hwKey || ! hwVal)
    {
        return E_OUTOFMEMORY;
    }
    static_cast<void>(yyjson_mut_obj_add(root, hwKey, hwVal));

    yyjson_mut_val* voutKey = yyjson_mut_strncpy(doc, "videoOutput", strlen("videoOutput"));
    yyjson_mut_val* voutVal = yyjson_mut_strncpy(doc, videoOutput.c_str(), videoOutput.size());
    if (! voutKey || ! voutVal)
    {
        return E_OUTOFMEMORY;
    }
    static_cast<void>(yyjson_mut_obj_add(root, voutKey, voutVal));

    yyjson_mut_val* aoutKey = yyjson_mut_strncpy(doc, "audioOutput", strlen("audioOutput"));
    yyjson_mut_val* aoutVal = yyjson_mut_strncpy(doc, audioOutput.c_str(), audioOutput.size());
    if (! aoutKey || ! aoutVal)
    {
        return E_OUTOFMEMORY;
    }
    static_cast<void>(yyjson_mut_obj_add(root, aoutKey, aoutVal));

    yyjson_mut_val* vizKey = yyjson_mut_strncpy(doc, "audioVisualization", strlen("audioVisualization"));
    yyjson_mut_val* vizVal = yyjson_mut_strncpy(doc, audioVisualization.c_str(), audioVisualization.size());
    if (! vizKey || ! vizVal)
    {
        return E_OUTOFMEMORY;
    }
    static_cast<void>(yyjson_mut_obj_add(root, vizKey, vizVal));

    yyjson_mut_val* extraKey = yyjson_mut_strncpy(doc, "extraArgs", strlen("extraArgs"));
    yyjson_mut_val* extraVal = yyjson_mut_strncpy(doc, extraArgs.c_str(), extraArgs.size());
    if (! extraKey || ! extraVal)
    {
        return E_OUTOFMEMORY;
    }
    static_cast<void>(yyjson_mut_obj_add(root, extraKey, extraVal));

    yyjson_write_err err{};
    size_t len = 0;
    wil::unique_any<char*, decltype(&::free), ::free> json(yyjson_mut_write_opts(doc, YYJSON_WRITE_NOFLAG, nullptr, &len, &err));
    if (! json || len == 0)
    {
        Debug::Warning(L"ViewerVLC: Failed to serialize configuration JSON: code: {}", err.code);
        _configurationJson = "{}";
        return S_OK;
    }

    _configurationJson.assign(json.get(), len);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerVLC::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *configurationJsonUtf8 = _configurationJson.empty() ? nullptr : _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerVLC::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    *pSomethingToSave = FALSE;
    return S_OK;
}

ATOM ViewerVLC::RegisterWndClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wc.lpfnWndProc   = &WndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kClassName;
    atom             = RegisterClassExW(&wc);
    if (atom == 0)
    {
        const DWORD lastError = GetLastError();
        if (lastError == ERROR_CLASS_ALREADY_EXISTS)
        {
            atom = 1;
        }
        else
        {
            Debug::ErrorWithLastError(L"ViewerVLC: RegisterClassExW failed.");
        }
    }
    return atom;
}

ATOM ViewerVLC::RegisterVideoClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wc.lpfnWndProc   = &VideoProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kVideoClassName;
    atom             = RegisterClassExW(&wc);
    if (atom == 0)
    {
        const DWORD lastError = GetLastError();
        if (lastError == ERROR_CLASS_ALREADY_EXISTS)
        {
            atom = 1;
        }
        else
        {
            Debug::ErrorWithLastError(L"ViewerVLC: RegisterVideoClass: RegisterClassExW failed.");
        }
    }
    return atom;
}

ATOM ViewerVLC::RegisterOverlayClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wc.lpfnWndProc   = &OverlayProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kOverlayClassName;
    atom             = RegisterClassExW(&wc);
    if (atom == 0)
    {
        const DWORD lastError = GetLastError();
        if (lastError == ERROR_CLASS_ALREADY_EXISTS)
        {
            atom = 1;
        }
        else
        {
            Debug::ErrorWithLastError(L"ViewerVLC: RegisterOverlayClass: RegisterClassExW failed.");
        }
    }
    return atom;
}

ATOM ViewerVLC::RegisterHudClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wc.lpfnWndProc   = &HudProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kHudClassName;
    atom             = RegisterClassExW(&wc);
    if (atom == 0)
    {
        const DWORD lastError = GetLastError();
        if (lastError == ERROR_CLASS_ALREADY_EXISTS)
        {
            atom = 1;
        }
        else
        {
            Debug::ErrorWithLastError(L"ViewerVLC: RegisterHudClass: RegisterClassExW failed.");
        }
    }
    return atom;
}

ATOM ViewerVLC::RegisterSeekPreviewClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = &SeekPreviewProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kSeekPreviewClassName;
    atom             = RegisterClassExW(&wc);
    if (atom == 0)
    {
        const DWORD lastError = GetLastError();
        if (lastError == ERROR_CLASS_ALREADY_EXISTS)
        {
            atom = 1;
        }
        else
        {
            Debug::ErrorWithLastError(L"ViewerVLC: RegisterSeekPreviewClass: RegisterClassExW failed.");
        }
    }
    return atom;
}

LRESULT CALLBACK ViewerVLC::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    ViewerVLC* self = reinterpret_cast<ViewerVLC*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! self && msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        self           = reinterpret_cast<ViewerVLC*>(cs ? cs->lpCreateParams : nullptr);
        if (self)
        {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
    }

    if (self)
    {
        return self->WndProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT CALLBACK ViewerVLC::VideoProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    ViewerVLC* self = reinterpret_cast<ViewerVLC*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! self && msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        self           = reinterpret_cast<ViewerVLC*>(cs ? cs->lpCreateParams : nullptr);
        if (self)
        {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
    }

    if (self)
    {
        return self->VideoProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT CALLBACK ViewerVLC::HudProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    ViewerVLC* self = reinterpret_cast<ViewerVLC*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! self && msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        self           = reinterpret_cast<ViewerVLC*>(cs ? cs->lpCreateParams : nullptr);
        if (self)
        {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
    }

    if (self)
    {
        return self->HudProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT CALLBACK ViewerVLC::OverlayProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    ViewerVLC* self = reinterpret_cast<ViewerVLC*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! self && msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        self           = reinterpret_cast<ViewerVLC*>(cs ? cs->lpCreateParams : nullptr);
        if (self)
        {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
    }

    if (self)
    {
        return self->OverlayProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT CALLBACK ViewerVLC::SeekPreviewProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    ViewerVLC* self = reinterpret_cast<ViewerVLC*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! self && msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        self           = reinterpret_cast<ViewerVLC*>(cs ? cs->lpCreateParams : nullptr);
        if (self)
        {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
    }

    if (self)
    {
        return self->SeekPreviewProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT ViewerVLC::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WM_SIZE: OnSize(LOWORD(lp), HIWORD(lp)); return 0;
        case WM_SETFOCUS:
            if (_hVideo)
            {
                SetFocus(_hVideo.get());
            }
            return 0;
        case WM_TIMER: OnTimer(static_cast<UINT_PTR>(wp)); return 0;
        case WM_KEYDOWN:
            if (wp == VK_ESCAPE)
            {
                if (_isFullscreen)
                {
                    SetFullscreen(false);
                }
                else
                {
                    DestroyWindow(hwnd);
                }
                return 0;
            }
            break;
        case WM_NCACTIVATE: ApplyTitleBarTheme(wp != FALSE); return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_LBUTTONDBLCLK: ToggleFullscreen(); return 0;
        case WM_PAINT:
        {
            PAINTSTRUCT ps{};
            wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
            if (hdc)
            {
                HBRUSH brush = _backgroundBrush ? _backgroundBrush.get() : GetSysColorBrush(COLOR_WINDOW);
                FillRect(hdc.get(), &ps.rcPaint, brush);
            }
            return 0;
        }
        case WM_ERASEBKGND: return 1;
        case WM_CLOSE: DestroyWindow(hwnd); return 0;
        case WM_NCDESTROY: return OnNcDestroy(hwnd, wp, lp);
        default: break;
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT ViewerVLC::VideoProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_LBUTTONDOWN:
            _hudLastActivityTick = 0;
            if (_hHud)
            {
                UpdateHudOpacityTarget(_hHud.get(), true);
            }
            SetFocus(hwnd);
            return 0;
        case WM_LBUTTONDBLCLK: ToggleFullscreen(); return 0;
        case WM_PARENTNOTIFY:
        {
            const UINT childMsg = LOWORD(wp);
            if (childMsg == WM_LBUTTONDBLCLK)
            {
                ToggleFullscreen();
                return 0;
            }
            if (childMsg == WM_LBUTTONDOWN)
            {
                const POINT pt      = {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))};
                const ULONGLONG now = GetTickCount64();
                const UINT maxDelay = GetDoubleClickTime();
                const int dx        = (pt.x >= _videoLastClickPos.x) ? (pt.x - _videoLastClickPos.x) : (_videoLastClickPos.x - pt.x);
                const int dy        = (pt.y >= _videoLastClickPos.y) ? (pt.y - _videoLastClickPos.y) : (_videoLastClickPos.y - pt.y);
                const int maxDx     = GetSystemMetrics(SM_CXDOUBLECLK);
                const int maxDy     = GetSystemMetrics(SM_CYDOUBLECLK);

                if (_videoLastClickTick != 0 && now >= _videoLastClickTick && (now - _videoLastClickTick) <= maxDelay && dx <= maxDx && dy <= maxDy)
                {
                    _videoLastClickTick = 0;
                    ToggleFullscreen();
                    return 0;
                }

                _videoLastClickTick = now;
                _videoLastClickPos  = pt;
            }
            break;
        }
        case WM_KEYDOWN:
        {
            const UINT vkey = static_cast<UINT>(wp);

            if (vkey == VK_ESCAPE)
            {
                if (_isFullscreen)
                {
                    SetFullscreen(false);
                }
                else if (_hWnd)
                {
                    _hWnd.reset();
                }
                return 0;
            }

            if ((GetKeyState(VK_MENU) & 0x8000) != 0)
            {
                break;
            }

            const bool ctrl  = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
            const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

            const auto markActivity = [&]
            {
                _hudLastActivityTick = GetTickCount64();
                if (_hHud)
                {
                    UpdateHudOpacityTarget(_hHud.get(), true);
                    InvalidateRect(_hHud.get(), nullptr, FALSE);
                }
            };

            if (vkey == VK_TAB)
            {
                if (_hHud && IsWindowVisible(_hHud.get()) != 0)
                {
                    SetFocus(_hHud.get());
                    markActivity();
                    return 0;
                }
                break;
            }

            if (vkey == VK_RETURN || vkey == VK_SPACE)
            {
                TogglePlayPause();
                markActivity();
                return 0;
            }

            if (! ctrl)
            {
                const int stepVolume = 5;
                switch (vkey)
                {
                    case VK_UP:
                        SetVolume(_hudVolumeValue + stepVolume);
                        markActivity();
                        return 0;
                    case VK_DOWN:
                        SetVolume(_hudVolumeValue - stepVolume);
                        markActivity();
                        return 0;
                    default: break;
                }

                const int64_t len       = _hudLengthMs;
                const int64_t stepSmall = (len > 0) ? std::max<int64_t>(1000, len / 200) : 5000;
                const int64_t stepLarge = (len > 0) ? std::max<int64_t>(5000, len / 20) : 30'000;
                const int64_t step      = shift ? stepLarge : stepSmall;

                switch (vkey)
                {
                    case VK_LEFT:
                        SeekRelativeMs(-step);
                        markActivity();
                        return 0;
                    case VK_RIGHT:
                        SeekRelativeMs(step);
                        markActivity();
                        return 0;
                    case VK_PRIOR:
                        SeekRelativeMs(-stepLarge);
                        markActivity();
                        return 0;
                    case VK_NEXT:
                        SeekRelativeMs(stepLarge);
                        markActivity();
                        return 0;
                    case VK_HOME:
                        SeekAbsoluteMs(0);
                        markActivity();
                        return 0;
                    case VK_END:
                        if (len > 0)
                        {
                            SeekAbsoluteMs(len);
                            markActivity();
                            return 0;
                        }
                        break;
                    default: break;
                }
            }
            break;
        }
        case WM_ERASEBKGND: return 1;
        case WM_PAINT:
        {
            PAINTSTRUCT ps{};
            wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
            if (hdc)
            {
                FillRect(hdc.get(), &ps.rcPaint, reinterpret_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
            }
            return 0;
        }
        case WM_NCDESTROY: SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0); return DefWindowProcW(hwnd, msg, wp, lp);
        default: break;
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT ViewerVLC::HudProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_GETDLGCODE: return DLGC_WANTARROWS | DLGC_WANTTAB | DLGC_WANTCHARS;
        case WM_SIZE: OnHudSize(hwnd, LOWORD(lp), HIWORD(lp)); return 0;
        case WM_TIMER:
            if (wp == kHudAnimTimerId)
            {
                OnHudTimer(hwnd);
                return 0;
            }
            break;
        case WM_MOUSEMOVE:
        {
            const POINT pt = {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))};
            OnHudMouseMove(hwnd, pt);
            return 0;
        }
        case WM_MOUSELEAVE: OnHudMouseLeave(hwnd); return 0;
        case WM_LBUTTONDOWN:
        {
            const POINT pt = {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))};
            OnHudLButtonDown(hwnd, pt);
            return 0;
        }
        case WM_LBUTTONUP:
        {
            const POINT pt = {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))};
            OnHudLButtonUp(hwnd, pt);
            return 0;
        }
        case WM_LBUTTONDBLCLK:
        {
            const POINT pt = {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))};
            if (HitTestHud(hwnd, pt) == HudPart::None)
            {
                ToggleFullscreen();
                return 0;
            }
            break;
        }
        case WM_KEYDOWN: OnHudKeyDown(hwnd, static_cast<UINT>(wp)); return 0;
        case WM_MOUSEWHEEL: OnHudMouseWheel(hwnd, static_cast<short>(GET_WHEEL_DELTA_WPARAM(wp))); return 0;
        case WM_SETFOCUS:
            _hudLastActivityTick = GetTickCount64();
            UpdateHudOpacityTarget(hwnd, true);
            return 0;
        case WM_KILLFOCUS: OnHudKillFocus(hwnd); return 0;
        case WM_PAINT: OnHudPaint(hwnd); return 0;
        case WM_ERASEBKGND: return 1;
        default: break;
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT ViewerVLC::OverlayProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_SIZE: OnOverlaySize(hwnd, LOWORD(lp), HIWORD(lp)); return 0;
        case WM_PAINT: OnOverlayPaint(hwnd); return 0;
        case WM_LBUTTONDBLCLK: ToggleFullscreen(); return 0;
        case WM_MOUSEMOVE:
        {
            const POINT pt = {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))};
            OnOverlayMouseMove(hwnd, pt);
            return 0;
        }
        case WM_MOUSELEAVE: OnOverlayMouseLeave(hwnd); return 0;
        case WM_LBUTTONUP:
        {
            const POINT pt = {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))};
            OnOverlayLButtonUp(hwnd, pt);
            return 0;
        }
        case WM_SETCURSOR:
        {
            const LRESULT handled = OnOverlaySetCursor(hwnd);
            if (handled != FALSE)
            {
                return handled;
            }
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        case WM_ERASEBKGND: return 1;
        default: return DefWindowProcW(hwnd, msg, wp, lp);
    }
}

LRESULT ViewerVLC::SeekPreviewProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_SIZE: OnSeekPreviewSize(hwnd, LOWORD(lp), HIWORD(lp)); return 0;
        case WM_PAINT: OnSeekPreviewPaint(hwnd); return 0;
        case WM_ERASEBKGND: return 1;
        case WM_NCHITTEST: return HTTRANSPARENT;
        case WM_NCDESTROY: SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0); return DefWindowProcW(hwnd, msg, wp, lp);
        default: return DefWindowProcW(hwnd, msg, wp, lp);
    }
}

void ViewerVLC::OnCreate(HWND hwnd) noexcept
{
    if (RegisterVideoClass(g_hInstance))
    {
        _hVideo.reset(CreateWindowExW(
            0, kVideoClassName, nullptr, WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, 0, 0, 0, 0, hwnd, nullptr, g_hInstance, this));
    }

    if (! _hVideo)
    {
        _hVideo.reset(CreateWindowExW(
            0, L"Static", nullptr, WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | SS_BLACKRECT, 0, 0, 0, 0, hwnd, nullptr, g_hInstance, nullptr));
    }

    if (RegisterHudClass(g_hInstance))
    {
        _hHud.reset(CreateWindowExW(0, kHudClassName, nullptr, WS_CHILD | WS_VISIBLE | WS_TABSTOP, 0, 0, 0, 0, hwnd, nullptr, g_hInstance, this));
    }

    if (RegisterOverlayClass(g_hInstance))
    {
        _hMissingOverlay.reset(CreateWindowExW(0, kOverlayClassName, nullptr, WS_CHILD | WS_CLIPSIBLINGS, 0, 0, 0, 0, hwnd, nullptr, g_hInstance, this));
    }

    if (RegisterSeekPreviewClass(g_hInstance))
    {
        _hSeekPreview.reset(CreateWindowExW(0, kSeekPreviewClassName, nullptr, WS_CHILD | WS_CLIPSIBLINGS, 0, 0, 0, 0, hwnd, nullptr, g_hInstance, this));
    }

    CreateOrUpdateWindowBackgroundBrush();
    ApplyTitleBarTheme(true);

    SetMissingUiVisible(false, {});
}

void ViewerVLC::OnDestroy() noexcept
{
    StopPlayback();

    IViewerCallback* callback = _callback;
    void* cookie              = _callbackCookie;
    if (callback)
    {
        AddRef();
        static_cast<void>(callback->ViewerClosed(cookie));
        Release();
    }
}

LRESULT ViewerVLC::OnNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp) noexcept
{
    OnDestroy();

    _hVideo.release();
    _hHud.release();
    _hMissingOverlay.release();
    _hSeekPreview.release();
    _hWnd.release();
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);

    _overlayLinkRect      = {};
    _overlayLinkHot       = false;
    _overlayTrackingMouse = false;
    _overlayDetails.clear();

    _backgroundBrush.reset();
    _backgroundColor = CLR_INVALID;

    _hudHot            = HudPart::None;
    _hudPressed        = HudPart::None;
    _hudFocus          = HudPart::PlayPause;
    _hudTrackingMouse  = false;
    _hudSeekDragging   = false;
    _hudVolumeDragging = false;
    _hudOpacity        = 1.0f;
    _hudTargetOpacity  = 1.0f;
    _hudAnimTimerId    = 0;
    _hudVolumeValue    = 100;
    _hudTimeMs         = 0;
    _hudLengthMs       = 0;
    _hudPlaying        = false;
    _hudDragTimeMs     = 0;

    _hudRenderTarget.reset();
    _hudTextFormat.reset();
    _hudMonoFormat.reset();

    _overlayRenderTarget.reset();
    _overlayTitleFormat.reset();
    _overlayBodyFormat.reset();
    _overlayLinkFormat.reset();
    _overlayTextDpi = 0;

    _hudDWriteFactory.reset();
    _hudD2DFactory.reset();

    Release();
    return DefWindowProcW(hwnd, WM_NCDESTROY, wp, lp);
}

void ViewerVLC::OnSize(UINT width, UINT height) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    Layout(_hWnd.get(), width, height);
}

HRESULT STDMETHODCALLTYPE ViewerVLC::Open(const ViewerOpenContext* context) noexcept
{
    if (! context || ! context->focusedPath || context->focusedPath[0] == L'\0')
    {
        Debug::Error(L"ViewerVLC: Open called with an invalid context (focusedPath missing).");
        return E_INVALIDARG;
    }

    const std::filesystem::path path(context->focusedPath);
    _currentPath = path;

    if (! _hWnd)
    {
        if (! RegisterWndClass(g_hInstance))
        {
            return E_FAIL;
        }

        HWND ownerWindow = context->ownerWindow;

        RECT ownerRect{};
        if (ownerWindow && GetWindowRect(ownerWindow, &ownerRect) != 0)
        {
            const int w = ownerRect.right - ownerRect.left;
            const int h = ownerRect.bottom - ownerRect.top;

            HWND window = CreateWindowExW(0,
                                          kClassName,
                                          LoadStringResource(g_hInstance, IDS_VIEWERVLC_WINDOW_CAPTION).c_str(),
                                          WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
                                          ownerRect.left,
                                          ownerRect.top,
                                          std::max(1, w),
                                          std::max(1, h),
                                          nullptr,
                                          nullptr,
                                          g_hInstance,
                                          this);
            if (! window)
            {
                const DWORD lastError = Debug::ErrorWithLastError(L"ViewerVLC: CreateWindowExW failed.");
                return HRESULT_FROM_WIN32(lastError);
            }

            _hWnd.reset(window);
            CreateOrUpdateWindowBackgroundBrush();
            ApplyTitleBarTheme(true);
        }
        else
        {
            HWND window = CreateWindowExW(0,
                                          kClassName,
                                          LoadStringResource(g_hInstance, IDS_VIEWERVLC_WINDOW_CAPTION).c_str(),
                                          WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
                                          CW_USEDEFAULT,
                                          CW_USEDEFAULT,
                                          900,
                                          700,
                                          nullptr,
                                          nullptr,
                                          g_hInstance,
                                          this);
            if (! window)
            {
                const DWORD lastError = Debug::ErrorWithLastError(L"ViewerVLC: CreateWindowExW failed.");
                return HRESULT_FROM_WIN32(lastError);
            }

            _hWnd.reset(window);
            CreateOrUpdateWindowBackgroundBrush();
            ApplyTitleBarTheme(true);
        }

        AddRef(); // Self-reference for window lifetime (released in WM_NCDESTROY)
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        static_cast<void>(SetForegroundWindow(_hWnd.get()));
    }
    else
    {
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        static_cast<void>(SetForegroundWindow(_hWnd.get()));
    }

    if (_hWnd)
    {
        SetWindowTextW(_hWnd.get(),
                       std::format(L"{} - {}", _currentPath.filename().wstring(), LoadStringResource(g_hInstance, IDS_VIEWERVLC_WINDOW_CAPTION)).c_str());
        ApplyTitleBarTheme(GetForegroundWindow() == _hWnd.get());
    }

    static_cast<void>(StartPlayback(_currentPath));
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerVLC::Close() noexcept
{
    _hWnd.reset();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerVLC::SetTheme(const ViewerTheme* theme) noexcept
{
    if (! theme || theme->version != 2)
    {
        return E_INVALIDARG;
    }

    _theme    = *theme;
    _hasTheme = true;

    if (_hWnd)
    {
        CreateOrUpdateWindowBackgroundBrush();
        ApplyTitleBarTheme(GetForegroundWindow() == _hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
    if (_hMissingOverlay)
    {
        InvalidateRect(_hMissingOverlay.get(), nullptr, TRUE);
    }
    if (_hHud)
    {
        InvalidateRect(_hHud.get(), nullptr, TRUE);
    }

    return S_OK;
}

void ViewerVLC::ApplyTitleBarTheme(bool windowActive) noexcept
{
    if (! _hasTheme || ! _hWnd)
    {
        return;
    }

    static constexpr DWORD kDwmwaUseImmersiveDarkMode19 = 19u;
    static constexpr DWORD kDwmwaUseImmersiveDarkMode20 = 20u;
    static constexpr DWORD kDwmwaBorderColor            = 34u;
    static constexpr DWORD kDwmwaCaptionColor           = 35u;
    static constexpr DWORD kDwmwaTextColor              = 36u;
    static constexpr DWORD kDwmColorDefault             = 0xFFFFFFFFu;

    const BOOL darkMode = (_theme.darkMode && ! _theme.highContrast) ? TRUE : FALSE;
    DwmSetWindowAttribute(_hWnd.get(), kDwmwaUseImmersiveDarkMode20, &darkMode, sizeof(darkMode));
    DwmSetWindowAttribute(_hWnd.get(), kDwmwaUseImmersiveDarkMode19, &darkMode, sizeof(darkMode));

    DWORD borderValue  = kDwmColorDefault;
    DWORD captionValue = kDwmColorDefault;
    DWORD textValue    = kDwmColorDefault;
    if (! _theme.highContrast && _theme.rainbowMode)
    {
        std::wstring_view seed = _currentPath.empty() ? std::wstring_view(L"title") : std::wstring_view(_currentPath.native());
        COLORREF accent        = ResolveAccentColor(_theme, seed);
        if (! windowActive)
        {
            static constexpr uint8_t kInactiveTitleBlendAlpha = 223u; // ~7/8 toward background
            const COLORREF bg                                 = ColorRefFromArgb(_theme.backgroundArgb);
            accent                                            = BlendColor(accent, bg, kInactiveTitleBlendAlpha);
        }

        const COLORREF text = ContrastingTextColor(accent);
        borderValue         = static_cast<DWORD>(accent);
        captionValue        = static_cast<DWORD>(accent);
        textValue           = static_cast<DWORD>(text);
    }

    DwmSetWindowAttribute(_hWnd.get(), kDwmwaBorderColor, &borderValue, sizeof(borderValue));
    DwmSetWindowAttribute(_hWnd.get(), kDwmwaCaptionColor, &captionValue, sizeof(captionValue));
    DwmSetWindowAttribute(_hWnd.get(), kDwmwaTextColor, &textValue, sizeof(textValue));
}

void ViewerVLC::CreateOrUpdateWindowBackgroundBrush() noexcept
{
    const COLORREF desired = (_hasTheme && ! _theme.highContrast) ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
    if (_backgroundBrush && _backgroundColor == desired)
    {
        return;
    }

    _backgroundColor = desired;
    _backgroundBrush.reset(CreateSolidBrush(desired));
    if (! _backgroundBrush)
    {
        _backgroundColor = CLR_INVALID;
    }
}

HRESULT STDMETHODCALLTYPE ViewerVLC::SetCallback(IViewerCallback* callback, void* cookie) noexcept
{
    _callback       = callback;
    _callbackCookie = cookie;
    return S_OK;
}

void ViewerVLC::TogglePlayPause() noexcept
{
    if (! _vlc || ! _vlc->player)
    {
        return;
    }

    const bool playing = _vlc->libvlc_media_player_is_playing && (_vlc->libvlc_media_player_is_playing(_vlc->player.get()) != 0);

    if (playing)
    {
        if (_vlc->libvlc_media_player_pause)
        {
            _vlc->libvlc_media_player_pause(_vlc->player.get());
        }
    }
    else
    {
        if (_vlc->libvlc_media_player_play)
        {
            static_cast<void>(_vlc->libvlc_media_player_play(_vlc->player.get()));
        }
    }

    UpdatePlaybackUi();
}

void ViewerVLC::StopCommand() noexcept
{
    if (_vlc && _vlc->player && _vlc->libvlc_media_player_stop)
    {
        _vlc->libvlc_media_player_stop(_vlc->player.get());
    }

    UpdatePlaybackUi();
}

void ViewerVLC::SeekAbsoluteMs(int64_t timeMs) noexcept
{
    if (! _vlc || ! _vlc->player || ! _vlc->libvlc_media_player_set_time)
    {
        return;
    }

    int64_t length = 0;
    if (_vlc->libvlc_media_player_get_length)
    {
        length = static_cast<int64_t>(_vlc->libvlc_media_player_get_length(_vlc->player.get()));
    }

    const int64_t clamped = (length > 0) ? std::clamp<int64_t>(timeMs, 0, length) : std::max<int64_t>(0, timeMs);
    _vlc->libvlc_media_player_set_time(_vlc->player.get(), static_cast<libvlc_time_t>(clamped));
    _hudTimeMs     = clamped;
    _hudDragTimeMs = clamped;
    UpdatePlaybackUi();
}

void ViewerVLC::SeekRelativeMs(int64_t deltaMs) noexcept
{
    SeekAbsoluteMs(_hudTimeMs + deltaMs);
}

void ViewerVLC::SetVolume(int volume) noexcept
{
    _hudVolumeValue = std::clamp(volume, 0, 100);
    if (_vlc && _vlc->player && _vlc->libvlc_audio_set_volume)
    {
        static_cast<void>(_vlc->libvlc_audio_set_volume(_vlc->player.get(), _hudVolumeValue));
    }
    if (_hHud)
    {
        UpdateHudOpacityTarget(_hHud.get(), false);
        InvalidateRect(_hHud.get(), nullptr, TRUE);
    }
}

void ViewerVLC::SetPlaybackRate(float rate) noexcept
{
    const float clamped = std::clamp(rate, 0.25f, 4.0f);
    _hudRate            = clamped;

    if (_vlc && _vlc->player && _vlc->libvlc_media_player_set_rate)
    {
        static_cast<void>(_vlc->libvlc_media_player_set_rate(_vlc->player.get(), clamped));
    }

    if (_hHud)
    {
        UpdateHudOpacityTarget(_hHud.get(), false);
        InvalidateRect(_hHud.get(), nullptr, TRUE);
    }
}

void ViewerVLC::StepPlaybackRate(int deltaSteps) noexcept
{
    static constexpr std::array<float, 7> kRates{{0.50f, 0.75f, 1.00f, 1.25f, 1.50f, 2.00f, 3.00f}};

    const float current = _hudRate;
    int bestIndex       = 2;
    float bestDelta     = std::numeric_limits<float>::max();
    for (int i = 0; i < static_cast<int>(kRates.size()); ++i)
    {
        const float d = std::fabs(kRates[static_cast<size_t>(i)] - current);
        if (d < bestDelta)
        {
            bestDelta = d;
            bestIndex = i;
        }
    }

    const int next = std::clamp(bestIndex + deltaSteps, 0, static_cast<int>(kRates.size()) - 1);
    SetPlaybackRate(kRates[static_cast<size_t>(next)]);
}

void ViewerVLC::ToggleFullscreen() noexcept
{
    SetFullscreen(! _isFullscreen);
}

void ViewerVLC::SetFullscreen(bool enabled) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    if (enabled == _isFullscreen)
    {
        return;
    }

    HWND hwnd = _hWnd.get();
    if (enabled)
    {
        _restoreStyle     = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_STYLE));
        _restoreExStyle   = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_EXSTYLE));
        _restorePlacement = WINDOWPLACEMENT{sizeof(WINDOWPLACEMENT)};
        static_cast<void>(GetWindowPlacement(hwnd, &_restorePlacement));

        MONITORINFO mi{sizeof(mi)};
        const HMONITOR monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
        if (GetMonitorInfoW(monitor, &mi) == 0)
        {
            return;
        }

        const DWORD newStyle   = (_restoreStyle & ~WS_OVERLAPPEDWINDOW) | WS_POPUP;
        const DWORD newExStyle = (_restoreExStyle & ~WS_EX_WINDOWEDGE);
        SetWindowLongPtrW(hwnd, GWL_STYLE, static_cast<LONG_PTR>(newStyle));
        SetWindowLongPtrW(hwnd, GWL_EXSTYLE, static_cast<LONG_PTR>(newExStyle));

        const RECT mrc = mi.rcMonitor;
        SetWindowPos(hwnd,
                     HWND_TOP,
                     mrc.left,
                     mrc.top,
                     std::max(1, static_cast<int>(mrc.right - mrc.left)),
                     std::max(1, static_cast<int>(mrc.bottom - mrc.top)),
                     SWP_NOOWNERZORDER | SWP_FRAMECHANGED);

        _isFullscreen = true;
    }
    else
    {
        SetWindowLongPtrW(hwnd, GWL_STYLE, static_cast<LONG_PTR>(_restoreStyle));
        SetWindowLongPtrW(hwnd, GWL_EXSTYLE, static_cast<LONG_PTR>(_restoreExStyle));

        if (_restorePlacement.length == sizeof(WINDOWPLACEMENT))
        {
            static_cast<void>(SetWindowPlacement(hwnd, &_restorePlacement));
        }

        SetWindowPos(hwnd, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOOWNERZORDER | SWP_FRAMECHANGED);
        _isFullscreen = false;
    }

    ApplyTitleBarTheme(GetForegroundWindow() == hwnd);
}

void ViewerVLC::OnTimer(UINT_PTR timerId) noexcept
{
    if (timerId != kUiTimerId)
    {
        return;
    }

    UpdatePlaybackUi();
}

LRESULT ViewerVLC::OnNotify(const NMHDR* hdr) noexcept
{
    if (! hdr)
    {
        return 0;
    }

    return 0;
}

void ViewerVLC::Layout(HWND hwnd, UINT width, UINT height) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const int w = static_cast<int>(width);
    const int h = static_cast<int>(height);
    if (w <= 0 || h <= 0)
    {
        return;
    }

    const UINT dpi   = GetDpiForWindow(hwnd);
    const auto scale = [&](int dip) noexcept { return MulDiv(dip, static_cast<int>(dpi), 96); };

    const int margin = scale(10);
    const int gap    = scale(8);

    const bool showHud  = ! _missingUiVisible;
    const int barHeight = showHud ? scale(64) : 0;

    const int barY = std::max(0, h - barHeight);

    if (_hVideo)
    {
        SetWindowPos(_hVideo.get(), nullptr, 0, 0, w, barY, SWP_NOZORDER | SWP_NOACTIVATE);
    }

    if (_hHud)
    {
        ShowWindow(_hHud.get(), showHud ? SW_SHOW : SW_HIDE);
        if (showHud)
        {
            SetWindowPos(_hHud.get(), HWND_TOP, 0, barY, w, barHeight, SWP_NOACTIVATE);
        }
    }

    if (_hMissingOverlay)
    {
        const int overlayMaxWidth = scale(680);
        const int overlayWidth    = std::min(std::max(1, w - 2 * margin), overlayMaxWidth);

        const int overlayPadding = scale(16);
        const int titleHeight    = scale(28);
        const int bodyHeight     = scale(92);
        const int linkHeight     = scale(22);
        const int overlayHeight  = overlayPadding * 2 + titleHeight + gap + bodyHeight + gap + linkHeight;

        const int overlayX = (w - overlayWidth) / 2;
        const int overlayY = std::max(margin, (barY - overlayHeight) / 3);

        SetWindowPos(_hMissingOverlay.get(), HWND_TOP, overlayX, overlayY, overlayWidth, overlayHeight, SWP_NOACTIVATE);
    }

    UpdateSeekPreviewLayout();
}

void ViewerVLC::UpdatePlaybackUi() noexcept
{
    if (_missingUiVisible)
    {
        return;
    }

    const bool hasPlayer = _vlc && _vlc->player != nullptr;
    bool playing         = false;
    int volume           = _hudVolumeValue;
    float rate           = _hudRate;
    int64_t nowMs        = 0;
    int64_t lenMs        = 0;

    if (hasPlayer)
    {
        if (_vlc->libvlc_media_player_is_playing)
        {
            playing = _vlc->libvlc_media_player_is_playing(_vlc->player.get()) != 0;
        }
        if (_vlc->libvlc_media_player_get_time)
        {
            nowMs = static_cast<int64_t>(_vlc->libvlc_media_player_get_time(_vlc->player.get()));
        }
        if (_vlc->libvlc_media_player_get_length)
        {
            lenMs = static_cast<int64_t>(_vlc->libvlc_media_player_get_length(_vlc->player.get()));
        }
        if (_vlc->libvlc_audio_get_volume)
        {
            const int v = _vlc->libvlc_audio_get_volume(_vlc->player.get());
            if (v >= 0)
            {
                volume = v;
            }
        }
        if (_vlc->libvlc_media_player_get_rate)
        {
            const float r = _vlc->libvlc_media_player_get_rate(_vlc->player.get());
            if (r > 0.0f)
            {
                rate = r;
            }
        }
    }

    _hudPlaying  = playing;
    _hudLengthMs = std::max<int64_t>(0, lenMs);
    if (! _hudSeekDragging)
    {
        _hudTimeMs     = std::max<int64_t>(0, nowMs);
        _hudDragTimeMs = _hudTimeMs;
    }
    _hudVolumeValue = std::clamp(volume, 0, 100);
    _hudRate        = std::clamp(rate, 0.25f, 4.0f);

    if (_hHud)
    {
        UpdateHudOpacityTarget(_hHud.get(), false);
        InvalidateRect(_hHud.get(), nullptr, TRUE);
    }
}

void ViewerVLC::SetMissingUiVisible(bool visible, std::wstring_view details) noexcept
{
    _missingUiVisible = visible;
    _overlayDetails.assign(details);
    _overlayLinkRect      = {};
    _overlayLinkHot       = false;
    _overlayTrackingMouse = false;

    if (_hMissingOverlay)
    {
        ShowWindow(_hMissingOverlay.get(), visible ? SW_SHOW : SW_HIDE);
        InvalidateRect(_hMissingOverlay.get(), nullptr, TRUE);
    }

    if (_hHud)
    {
        ShowWindow(_hHud.get(), visible ? SW_HIDE : SW_SHOW);
        InvalidateRect(_hHud.get(), nullptr, TRUE);
    }

    if (_hWnd)
    {
        RECT rc{};
        if (GetClientRect(_hWnd.get(), &rc) != 0)
        {
            Layout(_hWnd.get(), static_cast<UINT>(std::max<LONG>(0, rc.right - rc.left)), static_cast<UINT>(std::max<LONG>(0, rc.bottom - rc.top)));
        }
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }

    if (visible)
    {
        return;
    }

    UpdatePlaybackUi();
}

bool ViewerVLC::EnsureHudDirect2D(HWND hwnd) noexcept
{
    if (! _hudD2DFactory)
    {
        D2D1_FACTORY_OPTIONS options{};
        const HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, options, _hudD2DFactory.put());
        if (FAILED(hr))
        {
            return false;
        }
    }

    if (! _hudDWriteFactory)
    {
        const HRESULT hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(_hudDWriteFactory.put()));
        if (FAILED(hr))
        {
            return false;
        }
    }

    const UINT dpi = GetDpiForWindow(hwnd);

    if (! _hudRenderTarget)
    {
        RECT rc{};
        GetClientRect(hwnd, &rc);
        const UINT32 width  = static_cast<UINT32>(std::max<LONG>(0, rc.right - rc.left));
        const UINT32 height = static_cast<UINT32>(std::max<LONG>(0, rc.bottom - rc.top));

        D2D1_RENDER_TARGET_PROPERTIES rtProps = D2D1::RenderTargetProperties();
        rtProps.dpiX                          = 96.0f;
        rtProps.dpiY                          = 96.0f;

        const D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProps = D2D1::HwndRenderTargetProperties(hwnd, D2D1::SizeU(width, height));
        const HRESULT hr                                   = _hudD2DFactory->CreateHwndRenderTarget(rtProps, hwndProps, _hudRenderTarget.put());
        if (FAILED(hr))
        {
            return false;
        }

        _hudRenderTarget->SetDpi(96.0f, 96.0f);
        _hudRenderTarget->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
    }

    if (_hudTextDpi != dpi)
    {
        _hudTextDpi = dpi;
        _hudTextFormat.reset();
        _hudMonoFormat.reset();
    }

    if (! _hudTextFormat && _hudDWriteFactory)
    {
        const float size = static_cast<float>(MulDiv(12, static_cast<int>(dpi), 96));
        const HRESULT hr = _hudDWriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, size, L"", _hudTextFormat.put());
        if (SUCCEEDED(hr))
        {
            _hudTextFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _hudTextFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            _hudTextFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
        }
    }

    if (! _hudMonoFormat && _hudDWriteFactory)
    {
        const float size = static_cast<float>(MulDiv(12, static_cast<int>(dpi), 96));
        const HRESULT hr = _hudDWriteFactory->CreateTextFormat(
            L"Consolas", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, size, L"", _hudMonoFormat.put());
        if (SUCCEEDED(hr))
        {
            _hudMonoFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _hudMonoFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            _hudMonoFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING);
        }
    }

    return _hudRenderTarget != nullptr;
}

void ViewerVLC::DiscardHudRenderTarget() noexcept
{
    _hudRenderTarget.reset();
}

bool ViewerVLC::EnsureOverlayDirect2D(HWND hwnd) noexcept
{
    if (! _hudD2DFactory)
    {
        D2D1_FACTORY_OPTIONS options{};
        const HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, options, _hudD2DFactory.put());
        if (FAILED(hr))
        {
            return false;
        }
    }

    if (! _hudDWriteFactory)
    {
        const HRESULT hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(_hudDWriteFactory.put()));
        if (FAILED(hr))
        {
            return false;
        }
    }

    const UINT dpi = GetDpiForWindow(hwnd);

    if (! _overlayRenderTarget)
    {
        RECT rc{};
        GetClientRect(hwnd, &rc);
        const UINT32 width  = static_cast<UINT32>(std::max<LONG>(0, rc.right - rc.left));
        const UINT32 height = static_cast<UINT32>(std::max<LONG>(0, rc.bottom - rc.top));

        D2D1_RENDER_TARGET_PROPERTIES rtProps = D2D1::RenderTargetProperties();
        rtProps.dpiX                          = 96.0f;
        rtProps.dpiY                          = 96.0f;

        const D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProps = D2D1::HwndRenderTargetProperties(hwnd, D2D1::SizeU(width, height));
        const HRESULT hr                                   = _hudD2DFactory->CreateHwndRenderTarget(rtProps, hwndProps, _overlayRenderTarget.put());
        if (FAILED(hr))
        {
            return false;
        }

        _overlayRenderTarget->SetDpi(96.0f, 96.0f);
        _overlayRenderTarget->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
    }

    if (_overlayTextDpi != dpi)
    {
        _overlayTextDpi = dpi;
        _overlayTitleFormat.reset();
        _overlayBodyFormat.reset();
        _overlayLinkFormat.reset();
    }

    if (! _overlayTitleFormat && _hudDWriteFactory)
    {
        const float size = static_cast<float>(MulDiv(17, static_cast<int>(dpi), 96));
        const HRESULT hr = _hudDWriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_SEMI_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, size, L"", _overlayTitleFormat.put());
        if (SUCCEEDED(hr))
        {
            _overlayTitleFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);
            _overlayTitleFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
            _overlayTitleFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        }
    }

    if (! _overlayBodyFormat && _hudDWriteFactory)
    {
        const float size = static_cast<float>(MulDiv(12, static_cast<int>(dpi), 96));
        const HRESULT hr = _hudDWriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, size, L"", _overlayBodyFormat.put());
        if (SUCCEEDED(hr))
        {
            _overlayBodyFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);
            _overlayBodyFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
            _overlayBodyFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        }
    }

    if (! _overlayLinkFormat && _hudDWriteFactory)
    {
        const float size = static_cast<float>(MulDiv(12, static_cast<int>(dpi), 96));
        const HRESULT hr = _hudDWriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, size, L"", _overlayLinkFormat.put());
        if (SUCCEEDED(hr))
        {
            _overlayLinkFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _overlayLinkFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
            _overlayLinkFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        }
    }

    return _overlayRenderTarget != nullptr;
}

void ViewerVLC::DiscardOverlayRenderTarget() noexcept
{
    _overlayRenderTarget.reset();
}

void ViewerVLC::OnOverlaySize(HWND /*hwnd*/, UINT width, UINT height) noexcept
{
    if (! _overlayRenderTarget)
    {
        return;
    }

    const UINT32 w = static_cast<UINT32>(width);
    const UINT32 h = static_cast<UINT32>(height);
    _overlayRenderTarget->Resize(D2D1::SizeU(w, h));
}

bool ViewerVLC::EnsureSeekPreviewDirect2D(HWND hwnd) noexcept
{
    if (! _hudD2DFactory)
    {
        D2D1_FACTORY_OPTIONS options{};
        const HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, options, _hudD2DFactory.put());
        if (FAILED(hr))
        {
            return false;
        }
    }

    if (! _hudDWriteFactory)
    {
        const HRESULT hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(_hudDWriteFactory.put()));
        if (FAILED(hr))
        {
            return false;
        }
    }

    const UINT dpi = GetDpiForWindow(hwnd);

    if (! _seekPreviewRenderTarget)
    {
        RECT rc{};
        GetClientRect(hwnd, &rc);
        const UINT32 width  = static_cast<UINT32>(std::max<LONG>(0, rc.right - rc.left));
        const UINT32 height = static_cast<UINT32>(std::max<LONG>(0, rc.bottom - rc.top));

        D2D1_RENDER_TARGET_PROPERTIES rtProps = D2D1::RenderTargetProperties();
        rtProps.dpiX                          = 96.0f;
        rtProps.dpiY                          = 96.0f;

        const D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProps = D2D1::HwndRenderTargetProperties(hwnd, D2D1::SizeU(width, height));
        const HRESULT hr                                   = _hudD2DFactory->CreateHwndRenderTarget(rtProps, hwndProps, _seekPreviewRenderTarget.put());
        if (FAILED(hr))
        {
            return false;
        }

        _seekPreviewRenderTarget->SetDpi(96.0f, 96.0f);
        _seekPreviewRenderTarget->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
    }

    if (_seekPreviewTextDpi != dpi)
    {
        _seekPreviewTextDpi = dpi;
        _seekPreviewTextFormat.reset();
    }

    if (! _seekPreviewTextFormat && _hudDWriteFactory)
    {
        const float size = static_cast<float>(MulDiv(11, static_cast<int>(dpi), 96));
        const HRESULT hr = _hudDWriteFactory->CreateTextFormat(
            L"Consolas", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, size, L"", _seekPreviewTextFormat.put());
        if (SUCCEEDED(hr))
        {
            _seekPreviewTextFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _seekPreviewTextFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            _seekPreviewTextFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
        }
    }

    return _seekPreviewRenderTarget != nullptr;
}

void ViewerVLC::DiscardSeekPreviewRenderTarget() noexcept
{
    _seekPreviewRenderTarget.reset();
}

void ViewerVLC::OnSeekPreviewSize(HWND /*hwnd*/, UINT width, UINT height) noexcept
{
    if (! _seekPreviewRenderTarget)
    {
        return;
    }

    const UINT32 w = static_cast<UINT32>(width);
    const UINT32 h = static_cast<UINT32>(height);
    _seekPreviewRenderTarget->Resize(D2D1::SizeU(w, h));
}

void ViewerVLC::UpdateSeekPreviewTargetTimeMs(int64_t timeMs) noexcept
{
    if (! _hSeekPreview || _hudLengthMs <= 0)
    {
        return;
    }

    const int64_t clamped = std::clamp<int64_t>(timeMs, 0, _hudLengthMs);
    const int64_t quant   = (clamped / 1000) * 1000;
    if (quant == _seekPreviewTargetTimeMs)
    {
        return;
    }

    _seekPreviewTargetTimeMs = quant;
    InvalidateRect(_hSeekPreview.get(), nullptr, FALSE);
}

void ViewerVLC::UpdateSeekPreviewLayout() noexcept
{
    if (! _hWnd || ! _hHud || ! _hSeekPreview)
    {
        return;
    }

    const bool show = _hudSeekDragging && ! _missingUiVisible && (_hudLengthMs > 0) && (_vlc && _vlc->player);
    if (! show)
    {
        ShowWindow(_hSeekPreview.get(), SW_HIDE);
        return;
    }

    RECT hostRc{};
    if (GetClientRect(_hWnd.get(), &hostRc) == 0)
    {
        return;
    }
    const int hostW = std::max<LONG>(0, hostRc.right - hostRc.left);

    const UINT dpi = GetDpiForWindow(_hSeekPreview.get());
    const auto px  = [&](int dip) noexcept { return MulDiv(dip, static_cast<int>(dpi), 96); };

    const int w = px(112);
    const int h = px(34);

    POINT hudOrigin{0, 0};
    MapWindowPoints(_hHud.get(), _hWnd.get(), &hudOrigin, 1);

    RECT hudClient{};
    if (GetClientRect(_hHud.get(), &hudClient) == 0)
    {
        return;
    }

    const int hudW         = std::max<LONG>(0, hudClient.right - hudClient.left);
    const int hudH         = std::max<LONG>(0, hudClient.bottom - hudClient.top);
    const UINT hudDpi      = GetDpiForWindow(_hHud.get());
    const HudLayout layout = ComputeHudLayout(hudW, hudH, hudDpi);

    const int trackW   = std::max(1, static_cast<int>(layout.seekTrack.right - layout.seekTrack.left));
    const double ratio = (_hudLengthMs > 0) ? std::clamp(static_cast<double>(_hudDragTimeMs) / static_cast<double>(_hudLengthMs), 0.0, 1.0) : 0.0;
    const int fillW    = static_cast<int>(std::lround(static_cast<double>(trackW) * ratio));
    const int thumbX   = static_cast<int>(layout.seekTrack.left) + fillW;

    POINT thumbPt{thumbX, static_cast<int>(layout.seekTrack.top)};
    MapWindowPoints(_hHud.get(), _hWnd.get(), &thumbPt, 1);

    int x = thumbPt.x - (w / 2);
    int y = hudOrigin.y - h - px(8);

    const int clampMargin = px(8);
    x                     = std::clamp(x, clampMargin, std::max(clampMargin, hostW - w - clampMargin));
    y                     = std::max(clampMargin, y);

    SetWindowPos(_hSeekPreview.get(), HWND_TOP, x, y, w, h, SWP_NOACTIVATE | SWP_SHOWWINDOW);
}

void ViewerVLC::ClearSeekPreview() noexcept
{
    if (_hSeekPreview)
    {
        ShowWindow(_hSeekPreview.get(), SW_HIDE);
    }

    _seekPreviewTargetTimeMs = -1;
}

void ViewerVLC::OnSeekPreviewPaint(HWND hwnd) noexcept
{
    PAINTSTRUCT ps{};
    wil::unique_hdc_paint paintDc = wil::BeginPaint(hwnd, &ps);
    if (! paintDc)
    {
        return;
    }

    if (! EnsureSeekPreviewDirect2D(hwnd) || ! _seekPreviewRenderTarget)
    {
        return;
    }

    RECT client{};
    if (GetClientRect(hwnd, &client) == 0)
    {
        return;
    }

    const int w    = std::max<LONG>(0, client.right - client.left);
    const int h    = std::max<LONG>(0, client.bottom - client.top);
    const UINT dpi = GetDpiForWindow(hwnd);
    const auto px  = [&](int dip) noexcept { return MulDiv(dip, static_cast<int>(dpi), 96); };

    const bool themed            = _hasTheme && ! _theme.highContrast;
    const COLORREF windowBg      = themed ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
    const COLORREF windowFg      = themed ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);
    const std::wstring_view seed = _currentPath.empty() ? std::wstring_view(L"ViewerVLC") : std::wstring_view(_currentPath.native());
    const COLORREF accent        = themed ? ResolveAccentColor(_theme, seed) : GetSysColor(COLOR_HIGHLIGHT);

    const COLORREF cardBg = themed ? BlendColor(windowBg, windowFg, 18u) : GetSysColor(COLOR_WINDOW);
    const COLORREF border = themed ? BlendColor(cardBg, accent, 92u) : GetSysColor(COLOR_HIGHLIGHT);

    wil::com_ptr<ID2D1SolidColorBrush> brushBg;
    wil::com_ptr<ID2D1SolidColorBrush> brushBorder;
    wil::com_ptr<ID2D1SolidColorBrush> brushText;

    _seekPreviewRenderTarget->CreateSolidColorBrush(ColorFFromColorRef(cardBg), brushBg.put());
    _seekPreviewRenderTarget->CreateSolidColorBrush(ColorFFromColorRef(border), brushBorder.put());
    _seekPreviewRenderTarget->CreateSolidColorBrush(ColorFFromColorRef(windowFg), brushText.put());

    const float radius = static_cast<float>(px(10));
    const float stroke = std::max(1.0f, static_cast<float>(px(1)));

    const int padding        = px(6);
    const D2D1_RECT_F cardRc = D2D1::RectF(0.5f, 0.5f, static_cast<float>(w) - 0.5f, static_cast<float>(h) - 0.5f);
    const D2D1_RECT_F labelRc =
        D2D1::RectF(static_cast<float>(padding), static_cast<float>(padding), static_cast<float>(w - padding), static_cast<float>(h - padding));

    _seekPreviewRenderTarget->BeginDraw();
    _seekPreviewRenderTarget->Clear(ColorFFromColorRef(cardBg));

    _seekPreviewRenderTarget->FillRoundedRectangle(D2D1::RoundedRect(cardRc, radius, radius), brushBg.get());
    _seekPreviewRenderTarget->DrawRoundedRectangle(D2D1::RoundedRect(cardRc, radius, radius), brushBorder.get(), stroke);

    if (_seekPreviewTextFormat && _seekPreviewTargetTimeMs >= 0)
    {
        const std::wstring label = FormatDurationMs(static_cast<libvlc_time_t>(_seekPreviewTargetTimeMs));
        _seekPreviewRenderTarget->DrawTextW(
            label.c_str(), static_cast<UINT32>(label.size()), _seekPreviewTextFormat.get(), labelRc, brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
    }

    const HRESULT hr = _seekPreviewRenderTarget->EndDraw();
    if (hr == D2DERR_RECREATE_TARGET)
    {
        DiscardSeekPreviewRenderTarget();
    }
}

void ViewerVLC::UpdateHudOpacityTarget(HWND hwnd, bool forceInvalidate) noexcept
{
    if (_hasTheme && _theme.highContrast)
    {
        _hudOpacity       = 1.0f;
        _hudTargetOpacity = 1.0f;
        if (_hudAnimTimerId != 0)
        {
            KillTimer(hwnd, _hudAnimTimerId);
            _hudAnimTimerId = 0;
        }
        if (forceInvalidate)
        {
            InvalidateRect(hwnd, nullptr, FALSE);
        }
        return;
    }

    const bool hasPlayer      = _vlc && _vlc->player != nullptr;
    const ULONGLONG now       = GetTickCount64();
    const bool recentlyActive = (_hudLastActivityTick != 0) && (now >= _hudLastActivityTick) && ((now - _hudLastActivityTick) < kHudIdleDimDelayMs);
    const bool active         = _hudTrackingMouse || _hudSeekDragging || _hudVolumeDragging || (GetFocus() == hwnd) || recentlyActive;
    const float target        = (! hasPlayer) ? 1.0f : (active ? 1.0f : kHudDimOpacity);

    if (target == _hudTargetOpacity)
    {
        if (forceInvalidate)
        {
            InvalidateRect(hwnd, nullptr, FALSE);
        }
        return;
    }

    _hudTargetOpacity = target;
    if (_hudAnimTimerId == 0)
    {
        _hudAnimTimerId = SetTimer(hwnd, kHudAnimTimerId, kHudAnimIntervalMs, nullptr);
    }
    if (forceInvalidate)
    {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void ViewerVLC::CycleHudFocus(bool backwards) noexcept
{
    HudPart next = _hudFocus;
    switch (next)
    {
        case HudPart::None: next = HudPart::PlayPause; break;
        case HudPart::PlayPause: next = backwards ? HudPart::Volume : HudPart::Stop; break;
        case HudPart::Stop: next = backwards ? HudPart::PlayPause : HudPart::Snapshot; break;
        case HudPart::Snapshot: next = backwards ? HudPart::Stop : HudPart::Seek; break;
        case HudPart::Seek: next = backwards ? HudPart::Snapshot : HudPart::Speed; break;
        case HudPart::Speed: next = backwards ? HudPart::Seek : HudPart::Volume; break;
        case HudPart::Volume: next = backwards ? HudPart::Speed : HudPart::PlayPause; break;
    }
    _hudFocus = next;
}

ViewerVLC::HudPart ViewerVLC::HitTestHud(HWND hwnd, POINT pt) const noexcept
{
    RECT rc{};
    if (GetClientRect(hwnd, &rc) == 0)
    {
        return HudPart::None;
    }

    const int w            = std::max<LONG>(0, rc.right - rc.left);
    const int h            = std::max<LONG>(0, rc.bottom - rc.top);
    const UINT dpi         = GetDpiForWindow(hwnd);
    const HudLayout layout = ComputeHudLayout(w, h, dpi);

    if (PtInRect(&layout.play, pt) != 0)
        return HudPart::PlayPause;
    if (PtInRect(&layout.stop, pt) != 0)
        return HudPart::Stop;
    if (PtInRect(&layout.snapshot, pt) != 0)
        return HudPart::Snapshot;
    if (PtInRect(&layout.speed, pt) != 0)
        return HudPart::Speed;
    if (PtInRect(&layout.volume, pt) != 0)
        return HudPart::Volume;
    if (PtInRect(&layout.seekHit, pt) != 0)
        return HudPart::Seek;
    return HudPart::None;
}

void ViewerVLC::UpdateHudSeekDrag(HWND hwnd, POINT pt) noexcept
{
    if (! _hudSeekDragging || _hudLengthMs <= 0)
    {
        return;
    }

    RECT rc{};
    if (GetClientRect(hwnd, &rc) == 0)
    {
        return;
    }

    const int w            = std::max<LONG>(0, rc.right - rc.left);
    const int h            = std::max<LONG>(0, rc.bottom - rc.top);
    const UINT dpi         = GetDpiForWindow(hwnd);
    const HudLayout layout = ComputeHudLayout(w, h, dpi);

    const int trackW = std::max<int>(1, static_cast<int>(layout.seekTrack.right - layout.seekTrack.left));
    const int x      = pt.x - static_cast<int>(layout.seekTrack.left);
    const int px     = std::clamp(x, 0, trackW);
    const double t   = static_cast<double>(px) / static_cast<double>(trackW);
    _hudDragTimeMs   = static_cast<int64_t>(std::llround(static_cast<double>(_hudLengthMs) * t));

    UpdateSeekPreviewTargetTimeMs(_hudDragTimeMs);
    UpdateSeekPreviewLayout();
    InvalidateRect(hwnd, nullptr, FALSE);
}

void ViewerVLC::UpdateHudVolumeDrag(HWND hwnd, POINT pt) noexcept
{
    if (! _hudVolumeDragging)
    {
        return;
    }

    RECT rc{};
    if (GetClientRect(hwnd, &rc) == 0)
    {
        return;
    }

    const int w            = std::max<LONG>(0, rc.right - rc.left);
    const int h            = std::max<LONG>(0, rc.bottom - rc.top);
    const UINT dpi         = GetDpiForWindow(hwnd);
    const HudLayout layout = ComputeHudLayout(w, h, dpi);

    const int areaH = std::max<int>(1, static_cast<int>(layout.volume.bottom - layout.volume.top));
    const int y     = pt.y - static_cast<int>(layout.volume.top);
    const int px    = std::clamp(y, 0, areaH);
    const double t  = 1.0 - (static_cast<double>(px) / static_cast<double>(areaH));
    SetVolume(static_cast<int>(std::lround(t * 100.0)));
}

void ViewerVLC::OnHudSize(HWND /*hwnd*/, UINT width, UINT height) noexcept
{
    if (! _hudRenderTarget)
    {
        return;
    }

    const UINT32 w = static_cast<UINT32>(width);
    const UINT32 h = static_cast<UINT32>(height);
    _hudRenderTarget->Resize(D2D1::SizeU(w, h));
}

void ViewerVLC::OnHudTimer(HWND hwnd) noexcept
{
    const float diff = _hudTargetOpacity - _hudOpacity;
    if (std::fabs(diff) <= 0.02f)
    {
        _hudOpacity = _hudTargetOpacity;
        if (_hudAnimTimerId != 0)
        {
            KillTimer(hwnd, _hudAnimTimerId);
            _hudAnimTimerId = 0;
        }
    }
    else
    {
        _hudOpacity += diff * 0.25f;
    }

    InvalidateRect(hwnd, nullptr, FALSE);
}

void ViewerVLC::OnHudMouseMove(HWND hwnd, POINT pt) noexcept
{
    _hudLastActivityTick = GetTickCount64();

    if (! _hudTrackingMouse)
    {
        TRACKMOUSEEVENT tme{};
        tme.cbSize    = sizeof(tme);
        tme.dwFlags   = TME_LEAVE;
        tme.hwndTrack = hwnd;
        if (TrackMouseEvent(&tme) != 0)
        {
            _hudTrackingMouse = true;
            UpdateHudOpacityTarget(hwnd, false);
        }
    }

    if (_hudSeekDragging)
    {
        UpdateHudSeekDrag(hwnd, pt);
        return;
    }
    if (_hudVolumeDragging)
    {
        UpdateHudVolumeDrag(hwnd, pt);
        return;
    }

    const HudPart hot = HitTestHud(hwnd, pt);
    if (hot != _hudHot)
    {
        _hudHot = hot;
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void ViewerVLC::OnHudMouseLeave(HWND hwnd) noexcept
{
    _hudTrackingMouse = false;
    if (_hudHot != HudPart::None)
    {
        _hudHot = HudPart::None;
        InvalidateRect(hwnd, nullptr, FALSE);
    }
    UpdateHudOpacityTarget(hwnd, true);
}

void ViewerVLC::OnHudLButtonDown(HWND hwnd, POINT pt) noexcept
{
    SetFocus(hwnd);
    _hudLastActivityTick = GetTickCount64();

    const HudPart part = HitTestHud(hwnd, pt);
    _hudPressed        = part;
    if (part != HudPart::None)
    {
        _hudFocus = part;
    }

    UpdateHudOpacityTarget(hwnd, true);

    if (part == HudPart::Seek && _hudLengthMs > 0)
    {
        _seekDragWasPlaying = _hudPlaying;
        if (_seekDragWasPlaying && _vlc && _vlc->player && _vlc->libvlc_media_player_pause)
        {
            _vlc->libvlc_media_player_pause(_vlc->player.get());
        }
        _hudSeekDragging = true;
        SetCapture(hwnd);
        UpdateHudSeekDrag(hwnd, pt);
        return;
    }

    if (part == HudPart::Volume)
    {
        _hudVolumeDragging = true;
        SetCapture(hwnd);
        UpdateHudVolumeDrag(hwnd, pt);
        return;
    }

    if (part != HudPart::None)
    {
        SetCapture(hwnd);
    }

    InvalidateRect(hwnd, nullptr, FALSE);
}

void ViewerVLC::OnHudLButtonUp(HWND hwnd, POINT pt) noexcept
{
    _hudLastActivityTick = GetTickCount64();

    const HudPart part    = HitTestHud(hwnd, pt);
    const HudPart pressed = _hudPressed;
    _hudPressed           = HudPart::None;

    if (_hudSeekDragging)
    {
        _hudSeekDragging = false;
        if (GetCapture() == hwnd)
        {
            ReleaseCapture();
        }
        SeekAbsoluteMs(_hudDragTimeMs);
        if (_seekDragWasPlaying && _vlc && _vlc->player && _vlc->libvlc_media_player_play)
        {
            static_cast<void>(_vlc->libvlc_media_player_play(_vlc->player.get()));
        }
        _seekDragWasPlaying = false;
        ClearSeekPreview();
        UpdateHudOpacityTarget(hwnd, true);
        InvalidateRect(hwnd, nullptr, FALSE);
        return;
    }

    if (_hudVolumeDragging)
    {
        _hudVolumeDragging = false;
        if (GetCapture() == hwnd)
        {
            ReleaseCapture();
        }
        UpdateHudOpacityTarget(hwnd, true);
        InvalidateRect(hwnd, nullptr, FALSE);
        return;
    }

    if (GetCapture() == hwnd)
    {
        ReleaseCapture();
    }

    if (pressed != HudPart::None && pressed == part)
    {
        switch (pressed)
        {
            case HudPart::None: break;
            case HudPart::PlayPause: TogglePlayPause(); break;
            case HudPart::Stop: StopCommand(); break;
            case HudPart::Snapshot: TakeSnapshot(); break;
            case HudPart::Seek: break;
            case HudPart::Speed: StepPlaybackRate(1); break;
            case HudPart::Volume: break;
        }
    }

    UpdateHudOpacityTarget(hwnd, true);
    InvalidateRect(hwnd, nullptr, FALSE);
}

void ViewerVLC::OnHudKeyDown(HWND hwnd, UINT vkey) noexcept
{
    if (vkey == VK_ESCAPE)
    {
        if (_isFullscreen)
        {
            SetFullscreen(false);
            return;
        }

        if (_hWnd)
        {
            _hWnd.reset();
        }
        return;
    }

    _hudLastActivityTick = GetTickCount64();

    const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
    if (vkey == VK_TAB)
    {
        CycleHudFocus(shift);
        UpdateHudOpacityTarget(hwnd, true);
        return;
    }

    if (vkey == VK_RETURN || vkey == VK_SPACE)
    {
        switch (_hudFocus)
        {
            case HudPart::None: break;
            case HudPart::PlayPause: TogglePlayPause(); break;
            case HudPart::Stop: StopCommand(); break;
            case HudPart::Snapshot: TakeSnapshot(); break;
            case HudPart::Seek: break;
            case HudPart::Speed: StepPlaybackRate(1); break;
            case HudPart::Volume: break;
        }
        UpdateHudOpacityTarget(hwnd, true);
        return;
    }

    const int64_t len = _hudLengthMs;
    if (_hudFocus == HudPart::Seek && len > 0)
    {
        const int64_t stepSmall = std::max<int64_t>(1000, len / 200);
        const int64_t stepLarge = std::max<int64_t>(5000, len / 20);

        switch (vkey)
        {
            case VK_LEFT: SeekRelativeMs(-stepSmall); return;
            case VK_RIGHT: SeekRelativeMs(stepSmall); return;
            case VK_PRIOR: SeekRelativeMs(-stepLarge); return;
            case VK_NEXT: SeekRelativeMs(stepLarge); return;
            case VK_HOME: SeekAbsoluteMs(0); return;
            case VK_END: SeekAbsoluteMs(len); return;
            default: break;
        }
    }
    else if (_hudFocus == HudPart::Volume)
    {
        const int step = 5;
        switch (vkey)
        {
            case VK_UP: SetVolume(_hudVolumeValue + step); return;
            case VK_DOWN: SetVolume(_hudVolumeValue - step); return;
            case VK_LEFT: SetVolume(_hudVolumeValue - step); return;
            case VK_RIGHT: SetVolume(_hudVolumeValue + step); return;
            case VK_HOME: SetVolume(0); return;
            case VK_END: SetVolume(100); return;
            default: break;
        }
    }
    else if (_hudFocus == HudPart::Speed)
    {
        switch (vkey)
        {
            case VK_UP: StepPlaybackRate(1); return;
            case VK_DOWN: StepPlaybackRate(-1); return;
            case VK_LEFT: StepPlaybackRate(-1); return;
            case VK_RIGHT: StepPlaybackRate(1); return;
            case VK_PRIOR: StepPlaybackRate(1); return;
            case VK_NEXT: StepPlaybackRate(-1); return;
            case VK_HOME: SetPlaybackRate(1.0f); return;
            case VK_END: SetPlaybackRate(3.0f); return;
            default: break;
        }
    }
}

void ViewerVLC::OnHudMouseWheel(HWND /*hwnd*/, int wheelDelta) noexcept
{
    if (wheelDelta == 0)
    {
        return;
    }

    _hudLastActivityTick = GetTickCount64();

    if (_hudFocus == HudPart::Volume)
    {
        SetVolume(_hudVolumeValue + (wheelDelta > 0 ? 5 : -5));
        return;
    }

    if (_hudFocus == HudPart::Speed)
    {
        StepPlaybackRate(wheelDelta > 0 ? 1 : -1);
        return;
    }

    if (_hudFocus == HudPart::Seek && _hudLengthMs > 0)
    {
        SeekRelativeMs(wheelDelta > 0 ? 5000 : -5000);
        return;
    }
}

void ViewerVLC::OnHudKillFocus(HWND hwnd) noexcept
{
    _hudPressed        = HudPart::None;
    _hudSeekDragging   = false;
    _hudVolumeDragging = false;
    if (GetCapture() == hwnd)
    {
        ReleaseCapture();
    }
    UpdateHudOpacityTarget(hwnd, true);
    InvalidateRect(hwnd, nullptr, FALSE);
}

void ViewerVLC::OnHudPaint(HWND hwnd) noexcept
{
    PAINTSTRUCT ps{};
    wil::unique_hdc_paint paintDc = wil::BeginPaint(hwnd, &ps);
    if (! paintDc)
    {
        return;
    }

    if (! EnsureHudDirect2D(hwnd) || ! _hudRenderTarget)
    {
        return;
    }

    RECT client{};
    if (GetClientRect(hwnd, &client) == 0)
    {
        return;
    }

    const int w            = std::max<LONG>(0, client.right - client.left);
    const int h            = std::max<LONG>(0, client.bottom - client.top);
    const UINT dpi         = GetDpiForWindow(hwnd);
    const HudLayout layout = ComputeHudLayout(w, h, dpi);
    const auto px          = [&](int dip) noexcept { return MulDiv(dip, static_cast<int>(dpi), 96); };

    const bool themed            = _hasTheme && ! _theme.highContrast;
    const COLORREF bg            = themed ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
    const COLORREF fg            = themed ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);
    const std::wstring_view seed = _currentPath.empty() ? std::wstring_view(L"ViewerVLC") : std::wstring_view(_currentPath.native());
    const COLORREF accent        = themed ? ResolveAccentColor(_theme, seed) : GetSysColor(COLOR_HIGHLIGHT);

    const uint8_t dimA       = static_cast<uint8_t>(std::lround(std::clamp(_hudOpacity, 0.0f, 1.0f) * 255.0f));
    const COLORREF fgDim     = BlendColor(bg, fg, dimA);
    const COLORREF accentDim = BlendColor(bg, accent, dimA);

    const COLORREF border    = themed ? BlendColor(bg, fg, 64u) : GetSysColor(COLOR_WINDOWFRAME);
    const COLORREF borderDim = BlendColor(bg, border, dimA);

    const COLORREF hoverFill    = BlendColor(bg, accent, 34u);
    const COLORREF hoverFillDim = BlendColor(bg, hoverFill, dimA);
    const COLORREF pressFill    = BlendColor(bg, accent, 56u);
    const COLORREF pressFillDim = BlendColor(bg, pressFill, dimA);

    const uint8_t disabledA   = static_cast<uint8_t>(std::lround(static_cast<float>(dimA) * 0.55f));
    const COLORREF fgDisabled = BlendColor(bg, fg, disabledA);

    wil::com_ptr<ID2D1SolidColorBrush> brushText;
    wil::com_ptr<ID2D1SolidColorBrush> brushTextDisabled;
    wil::com_ptr<ID2D1SolidColorBrush> brushAccent;
    wil::com_ptr<ID2D1SolidColorBrush> brushBorder;
    wil::com_ptr<ID2D1SolidColorBrush> brushHover;
    wil::com_ptr<ID2D1SolidColorBrush> brushPress;

    _hudRenderTarget->CreateSolidColorBrush(ColorFFromColorRef(fgDim), brushText.put());
    _hudRenderTarget->CreateSolidColorBrush(ColorFFromColorRef(fgDisabled), brushTextDisabled.put());
    _hudRenderTarget->CreateSolidColorBrush(ColorFFromColorRef(accentDim), brushAccent.put());
    _hudRenderTarget->CreateSolidColorBrush(ColorFFromColorRef(borderDim), brushBorder.put());
    _hudRenderTarget->CreateSolidColorBrush(ColorFFromColorRef(hoverFillDim), brushHover.put());
    _hudRenderTarget->CreateSolidColorBrush(ColorFFromColorRef(pressFillDim), brushPress.put());

    const float radius = static_cast<float>(px(6));
    const float stroke = std::max(1.0f, static_cast<float>(px(1)));

    const bool hasPlayer = _vlc && _vlc->player != nullptr;
    const bool allowSeek = hasPlayer && _hudLengthMs > 0;

    const auto drawButtonBackground = [&](const RECT& rcBtn, HudPart part, bool enabled)
    {
        const bool hot     = (_hudHot == part);
        const bool pressed = (_hudPressed == part);
        const bool focused = (_hudFocus == part) && (GetFocus() == hwnd);

        if (enabled && (hot || pressed || focused))
        {
            const auto fillBrush = pressed ? brushPress.get() : brushHover.get();
            _hudRenderTarget->FillRoundedRectangle(D2D1::RoundedRect(RectFFromRect(rcBtn), radius, radius), fillBrush);
        }

        _hudRenderTarget->DrawRoundedRectangle(D2D1::RoundedRect(RectFFromRect(rcBtn), radius, radius), brushBorder.get(), stroke);

        if (focused)
        {
            const float focusStroke = stroke * 2.0f;
            _hudRenderTarget->DrawRoundedRectangle(D2D1::RoundedRect(RectFFromRect(rcBtn), radius + 1.0f, radius + 1.0f), brushAccent.get(), focusStroke);
        }
    };

    const auto fillTriangle = [&](D2D1_POINT_2F p1, D2D1_POINT_2F p2, D2D1_POINT_2F p3, ID2D1Brush* brush)
    {
        if (! _hudD2DFactory)
        {
            return;
        }

        wil::com_ptr<ID2D1PathGeometry> geo;
        if (FAILED(_hudD2DFactory->CreatePathGeometry(geo.put())))
        {
            return;
        }

        wil::com_ptr<ID2D1GeometrySink> sink;
        if (FAILED(geo->Open(sink.put())))
        {
            return;
        }

        sink->BeginFigure(p1, D2D1_FIGURE_BEGIN_FILLED);
        std::array<D2D1_POINT_2F, 2> pts{p2, p3};
        sink->AddLines(pts.data(), static_cast<UINT32>(pts.size()));
        sink->EndFigure(D2D1_FIGURE_END_CLOSED);
        static_cast<void>(sink->Close());

        _hudRenderTarget->FillGeometry(geo.get(), brush);
    };

    const auto drawPlayPauseIcon = [&](const RECT& rcBtn, ID2D1Brush* brush)
    {
        const float cx = (static_cast<float>(rcBtn.left) + static_cast<float>(rcBtn.right)) * 0.5f;
        const float cy = (static_cast<float>(rcBtn.top) + static_cast<float>(rcBtn.bottom)) * 0.5f;
        const float s  = std::max(6.0f, (static_cast<float>(rcBtn.right - rcBtn.left) * 0.32f));

        if (_hudPlaying)
        {
            const float barW     = std::max(2.0f, s * 0.28f);
            const float barH     = s * 1.15f;
            const float gapX     = barW * 0.55f;
            const D2D1_RECT_F r1 = D2D1::RectF(cx - gapX - barW, cy - barH * 0.5f, cx - gapX, cy + barH * 0.5f);
            const D2D1_RECT_F r2 = D2D1::RectF(cx + gapX, cy - barH * 0.5f, cx + gapX + barW, cy + barH * 0.5f);
            _hudRenderTarget->FillRectangle(r1, brush);
            _hudRenderTarget->FillRectangle(r2, brush);
        }
        else
        {
            const D2D1_POINT_2F p1 = {cx - s * 0.55f, cy - s * 0.75f};
            const D2D1_POINT_2F p2 = {cx - s * 0.55f, cy + s * 0.75f};
            const D2D1_POINT_2F p3 = {cx + s * 0.80f, cy};
            fillTriangle(p1, p2, p3, brush);
        }
    };

    const auto drawStopIcon = [&](const RECT& rcBtn, ID2D1Brush* brush)
    {
        const float cx      = (static_cast<float>(rcBtn.left) + static_cast<float>(rcBtn.right)) * 0.5f;
        const float cy      = (static_cast<float>(rcBtn.top) + static_cast<float>(rcBtn.bottom)) * 0.5f;
        const float s       = std::max(6.0f, (static_cast<float>(rcBtn.right - rcBtn.left) * 0.34f));
        const D2D1_RECT_F r = D2D1::RectF(cx - s, cy - s, cx + s, cy + s);
        _hudRenderTarget->FillRoundedRectangle(D2D1::RoundedRect(r, radius * 0.5f, radius * 0.5f), brush);
    };

    const auto drawSnapshotIcon = [&](const RECT& rcBtn, ID2D1Brush* brush)
    {
        const float cx         = (static_cast<float>(rcBtn.left) + static_cast<float>(rcBtn.right)) * 0.5f;
        const float cy         = (static_cast<float>(rcBtn.top) + static_cast<float>(rcBtn.bottom)) * 0.5f;
        const float s          = std::max(6.0f, (static_cast<float>(rcBtn.right - rcBtn.left) * 0.34f));
        const D2D1_RECT_F body = D2D1::RectF(cx - s, cy - s * 0.55f, cx + s, cy + s * 0.65f);
        _hudRenderTarget->DrawRoundedRectangle(D2D1::RoundedRect(body, radius * 0.5f, radius * 0.5f), brush, stroke);
        const D2D1_ELLIPSE lens{D2D1::Point2F(cx, cy + s * 0.05f), s * 0.35f, s * 0.35f};
        _hudRenderTarget->DrawEllipse(lens, brush, stroke);
        const D2D1_RECT_F top = D2D1::RectF(cx - s * 0.55f, cy - s * 0.75f, cx - s * 0.05f, cy - s * 0.55f);
        _hudRenderTarget->FillRoundedRectangle(D2D1::RoundedRect(top, radius * 0.25f, radius * 0.25f), brush);
    };

    _hudRenderTarget->BeginDraw();
    _hudRenderTarget->Clear(ColorFFromColorRef(bg));

    drawButtonBackground(layout.play, HudPart::PlayPause, hasPlayer);
    drawPlayPauseIcon(layout.play, hasPlayer ? brushText.get() : brushTextDisabled.get());

    drawButtonBackground(layout.stop, HudPart::Stop, hasPlayer);
    drawStopIcon(layout.stop, hasPlayer ? brushText.get() : brushTextDisabled.get());

    drawButtonBackground(layout.snapshot, HudPart::Snapshot, hasPlayer);
    drawSnapshotIcon(layout.snapshot, hasPlayer ? brushText.get() : brushTextDisabled.get());

    // Seek bar
    const RECT trackRc      = layout.seekTrack;
    const float trackRadius = std::max(1.0f, static_cast<float>(px(3)));
    _hudRenderTarget->FillRoundedRectangle(D2D1::RoundedRect(RectFFromRect(trackRc), trackRadius, trackRadius), brushBorder.get());

    const int trackW    = std::max(1, static_cast<int>(trackRc.right - trackRc.left));
    const int64_t lenMs = _hudLengthMs;
    const int64_t posMs = _hudSeekDragging ? _hudDragTimeMs : _hudTimeMs;
    const double ratio  = (lenMs > 0) ? std::clamp(static_cast<double>(posMs) / static_cast<double>(lenMs), 0.0, 1.0) : 0.0;
    const int fillW     = static_cast<int>(std::lround(static_cast<double>(trackW) * ratio));
    if (fillW > 0)
    {
        RECT fillRc  = trackRc;
        fillRc.right = std::min(trackRc.right, trackRc.left + fillW);
        _hudRenderTarget->FillRoundedRectangle(D2D1::RoundedRect(RectFFromRect(fillRc), trackRadius, trackRadius),
                                               allowSeek ? brushAccent.get() : brushBorder.get());
    }

    const float thumbX = static_cast<float>(trackRc.left + fillW);
    const float thumbY = (static_cast<float>(trackRc.top) + static_cast<float>(trackRc.bottom)) * 0.5f;
    const float thumbR = std::max(3.0f, static_cast<float>(px(6)));
    const D2D1_ELLIPSE thumb{D2D1::Point2F(thumbX, thumbY), thumbR, thumbR};
    _hudRenderTarget->FillEllipse(thumb, allowSeek ? brushText.get() : brushTextDisabled.get());
    _hudRenderTarget->DrawEllipse(thumb, allowSeek ? brushAccent.get() : brushBorder.get(), stroke);

    const bool seekFocused = (_hudFocus == HudPart::Seek) && (GetFocus() == hwnd);
    if (seekFocused)
    {
        RECT focusRc = layout.seekHit;
        focusRc.top += px(6);
        focusRc.bottom -= px(6);
        _hudRenderTarget->DrawRoundedRectangle(D2D1::RoundedRect(RectFFromRect(focusRc), radius, radius), brushAccent.get(), stroke * 2.0f);
    }

    // Time label
    if (! IsRectEmpty(&layout.time) && _hudMonoFormat)
    {
        const std::wstring label =
            (lenMs > 0) ? std::format(L"{} / {}", FormatDurationMs(static_cast<libvlc_time_t>(posMs)), FormatDurationMs(static_cast<libvlc_time_t>(lenMs)))
                        : LoadStringResource(g_hInstance, IDS_VIEWERVLC_LABEL_TIME_UNKNOWN);

        const D2D1_RECT_F textRc = RectFFromRect(layout.time);
        _hudRenderTarget->DrawTextW(
            label.c_str(), static_cast<UINT32>(label.size()), _hudMonoFormat.get(), textRc, brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
    }

    // Speed
    if (! IsRectEmpty(&layout.speed))
    {
        drawButtonBackground(layout.speed, HudPart::Speed, hasPlayer);
        if (_hudTextFormat)
        {
            const std::wstring rateText = FormatPlaybackRate(_hudRate);
            const D2D1_RECT_F rateRc    = RectFFromRect(layout.speed);
            _hudRenderTarget->DrawTextW(rateText.c_str(),
                                        static_cast<UINT32>(rateText.size()),
                                        _hudTextFormat.get(),
                                        rateRc,
                                        hasPlayer ? brushText.get() : brushTextDisabled.get(),
                                        D2D1_DRAW_TEXT_OPTIONS_CLIP);
        }
    }

    // Volume knob
    drawButtonBackground(layout.volume, HudPart::Volume, true);
    const float vcx = (static_cast<float>(layout.volume.left) + static_cast<float>(layout.volume.right)) * 0.5f;
    const float vcy = (static_cast<float>(layout.volume.top) + static_cast<float>(layout.volume.bottom)) * 0.5f;
    const float vr  = std::max(6.0f, (static_cast<float>(layout.volume.right - layout.volume.left) * 0.32f));
    const D2D1_ELLIPSE vknob{D2D1::Point2F(vcx, vcy), vr, vr};
    _hudRenderTarget->DrawEllipse(vknob, brushBorder.get(), stroke);
    _hudRenderTarget->DrawEllipse(vknob, brushAccent.get(), stroke);

    const double vRatio = std::clamp(static_cast<double>(_hudVolumeValue) / 100.0, 0.0, 1.0);
    const double startA = -3.14159265358979323846 * 0.75;
    const double sweep  = 3.14159265358979323846 * 1.5;
    const double ang    = startA + sweep * vRatio;
    const float ix      = vcx + static_cast<float>(std::cos(ang) * (vr * 0.85));
    const float iy      = vcy + static_cast<float>(std::sin(ang) * (vr * 0.85));
    _hudRenderTarget->DrawLine(D2D1::Point2F(vcx, vcy), D2D1::Point2F(ix, iy), brushText.get(), stroke * 1.5f);

    const bool volFocused = (_hudFocus == HudPart::Volume) && (GetFocus() == hwnd);
    if (((_hudHot == HudPart::Volume) || volFocused) && _hudTextFormat)
    {
        const std::wstring vText = std::format(L"{}%", _hudVolumeValue);
        const D2D1_RECT_F vRc    = RectFFromRect(layout.volume);
        _hudRenderTarget->DrawTextW(vText.c_str(), static_cast<UINT32>(vText.size()), _hudTextFormat.get(), vRc, brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
    }

    // Top border line
    _hudRenderTarget->DrawLine(D2D1::Point2F(0.0f, 0.5f), D2D1::Point2F(static_cast<float>(w), 0.5f), brushBorder.get(), stroke);

    const HRESULT hr = _hudRenderTarget->EndDraw();
    if (hr == D2DERR_RECREATE_TARGET)
    {
        DiscardHudRenderTarget();
    }
}

void ViewerVLC::OnOverlayPaint(HWND hwnd) noexcept
{
    PAINTSTRUCT ps{};
    wil::unique_hdc_paint paintDc = wil::BeginPaint(hwnd, &ps);
    if (! paintDc)
    {
        return;
    }

    if (! EnsureOverlayDirect2D(hwnd) || ! _overlayRenderTarget)
    {
        return;
    }

    RECT rc{};
    if (GetClientRect(hwnd, &rc) == 0)
    {
        return;
    }

    const int w = std::max<LONG>(0, rc.right - rc.left);
    const int h = std::max<LONG>(0, rc.bottom - rc.top);
    if (w <= 0 || h <= 0)
    {
        return;
    }

    const UINT dpi = GetDpiForWindow(hwnd);
    const auto px  = [&](int dip) noexcept { return MulDiv(dip, static_cast<int>(dpi), 96); };

    const bool themed = _hasTheme && ! _theme.highContrast;

    const COLORREF cardBg = themed ? ColorRefFromArgb(_theme.alertInfoBackgroundArgb) : GetSysColor(COLOR_INFOBK);
    const COLORREF cardFg = themed ? ColorRefFromArgb(_theme.alertInfoTextArgb) : GetSysColor(COLOR_INFOTEXT);

    const std::wstring_view seed = _currentPath.empty() ? std::wstring_view(L"ViewerVLC") : std::wstring_view(_currentPath.native());
    const COLORREF accent        = themed ? ResolveAccentColor(_theme, seed) : GetSysColor(COLOR_HIGHLIGHT);
    const COLORREF border        = themed ? BlendColor(cardBg, accent, 92u) : GetSysColor(COLOR_HIGHLIGHT);

    const COLORREF linkFg    = themed ? accent : GetSysColor(COLOR_HOTLIGHT);
    const COLORREF linkFgHot = themed ? BlendColor(linkFg, cardFg, 96u) : GetSysColor(COLOR_HIGHLIGHT);

    const int stripeW = px(6);
    const int padding = px(16);
    const int gap     = px(8);

    const float radius = static_cast<float>(px(8));
    const float stroke = std::max(1.0f, static_cast<float>(px(1)));

    RECT contentRc{padding + stripeW, padding, w - padding, h - padding};
    if (contentRc.right <= contentRc.left || contentRc.bottom <= contentRc.top)
    {
        return;
    }

    const std::wstring title     = GetOverlayTitleText();
    const std::wstring body      = GetOverlayBodyText();
    const std::wstring linkLabel = GetOverlayLinkLabelText();

    RECT titleRc   = contentRc;
    titleRc.bottom = titleRc.top;
    int y          = contentRc.top;

    if (! title.empty() && _hudDWriteFactory && _overlayTitleFormat)
    {
        wil::com_ptr<IDWriteTextLayout> titleLayout;
        const float layoutW = static_cast<float>(contentRc.right - contentRc.left);
        const float layoutH = static_cast<float>(contentRc.bottom - contentRc.top);
        if (SUCCEEDED(_hudDWriteFactory->CreateTextLayout(
                title.c_str(), static_cast<UINT32>(title.size()), _overlayTitleFormat.get(), layoutW, layoutH, titleLayout.put())))
        {
            DWRITE_TEXT_METRICS metrics{};
            if (SUCCEEDED(titleLayout->GetMetrics(&metrics)))
            {
                const int titleH = std::max(0, static_cast<int>(std::ceil(metrics.height)));
                titleRc.bottom   = std::min(contentRc.bottom, titleRc.top + titleH);
                y                = std::min(contentRc.bottom, titleRc.bottom + gap);
            }
        }
    }

    RECT linkRc{};
    _overlayLinkRect = {};
    wil::com_ptr<IDWriteTextLayout> linkLayout;

    if (! linkLabel.empty() && _hudDWriteFactory && _overlayLinkFormat)
    {
        const float layoutW = static_cast<float>(contentRc.right - contentRc.left);
        const float layoutH = static_cast<float>(contentRc.bottom - contentRc.top);
        if (SUCCEEDED(_hudDWriteFactory->CreateTextLayout(
                linkLabel.c_str(), static_cast<UINT32>(linkLabel.size()), _overlayLinkFormat.get(), layoutW, layoutH, linkLayout.put())))
        {
            DWRITE_TEXT_METRICS metrics{};
            if (SUCCEEDED(linkLayout->GetMetrics(&metrics)))
            {
                const int linkW = std::max(0, static_cast<int>(std::ceil(metrics.widthIncludingTrailingWhitespace)));
                const int linkH = std::max(0, static_cast<int>(std::ceil(metrics.height)));

                linkRc.left   = contentRc.left;
                linkRc.right  = std::min(contentRc.right, contentRc.left + linkW);
                linkRc.bottom = contentRc.bottom;
                linkRc.top    = std::max(y, static_cast<int>(linkRc.bottom) - linkH);

                if (! IsRectEmpty(&linkRc))
                {
                    _overlayLinkRect = linkRc;
                }
            }

            const DWRITE_TEXT_RANGE full{0, static_cast<UINT32>(linkLabel.size())};
            static_cast<void>(linkLayout->SetUnderline(TRUE, full));
        }
    }

    RECT bodyRc = contentRc;
    bodyRc.top  = y;
    if (! IsRectEmpty(&linkRc))
    {
        bodyRc.bottom = std::max(bodyRc.top, linkRc.top - gap);
    }

    wil::com_ptr<ID2D1SolidColorBrush> brushText;
    wil::com_ptr<ID2D1SolidColorBrush> brushBorder;
    wil::com_ptr<ID2D1SolidColorBrush> brushAccent;
    wil::com_ptr<ID2D1SolidColorBrush> brushLink;

    _overlayRenderTarget->CreateSolidColorBrush(ColorFFromColorRef(cardFg), brushText.put());
    _overlayRenderTarget->CreateSolidColorBrush(ColorFFromColorRef(border), brushBorder.put());
    _overlayRenderTarget->CreateSolidColorBrush(ColorFFromColorRef(accent), brushAccent.put());
    _overlayRenderTarget->CreateSolidColorBrush(ColorFFromColorRef(_overlayLinkHot ? linkFgHot : linkFg), brushLink.put());

    _overlayRenderTarget->BeginDraw();
    _overlayRenderTarget->Clear(ColorFFromColorRef(cardBg));

    const D2D1_RECT_F cardRc = D2D1::RectF(0.5f, 0.5f, static_cast<float>(w) - 0.5f, static_cast<float>(h) - 0.5f);
    _overlayRenderTarget->DrawRoundedRectangle(D2D1::RoundedRect(cardRc, radius, radius), brushBorder.get(), stroke);

    const D2D1_RECT_F stripeRc = D2D1::RectF(1.0f, 1.0f, static_cast<float>(1 + stripeW), static_cast<float>(h) - 1.0f);
    _overlayRenderTarget->FillRectangle(stripeRc, brushAccent.get());

    if (titleRc.bottom > titleRc.top && _overlayTitleFormat)
    {
        _overlayRenderTarget->DrawTextW(
            title.c_str(), static_cast<UINT32>(title.size()), _overlayTitleFormat.get(), RectFFromRect(titleRc), brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
    }

    if (! body.empty() && bodyRc.bottom > bodyRc.top && _overlayBodyFormat)
    {
        _overlayRenderTarget->DrawTextW(
            body.c_str(), static_cast<UINT32>(body.size()), _overlayBodyFormat.get(), RectFFromRect(bodyRc), brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
    }

    if (! IsRectEmpty(&linkRc) && linkLayout)
    {
        const D2D1_POINT_2F origin{static_cast<float>(linkRc.left), static_cast<float>(linkRc.top)};
        _overlayRenderTarget->DrawTextLayout(origin, linkLayout.get(), brushLink.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
    }

    const HRESULT hr = _overlayRenderTarget->EndDraw();
    if (hr == D2DERR_RECREATE_TARGET)
    {
        DiscardOverlayRenderTarget();
    }
}

void ViewerVLC::OnOverlayMouseMove(HWND hwnd, POINT pt) noexcept
{
    const bool hot = PtInRect(&_overlayLinkRect, pt) != 0;
    if (hot != _overlayLinkHot)
    {
        _overlayLinkHot = hot;
        InvalidateRect(hwnd, nullptr, TRUE);
    }

    if (! _overlayTrackingMouse)
    {
        TRACKMOUSEEVENT tme{};
        tme.cbSize    = sizeof(tme);
        tme.dwFlags   = TME_LEAVE;
        tme.hwndTrack = hwnd;
        if (TrackMouseEvent(&tme) != 0)
        {
            _overlayTrackingMouse = true;
        }
    }
}

void ViewerVLC::OnOverlayMouseLeave(HWND hwnd) noexcept
{
    _overlayTrackingMouse = false;
    if (_overlayLinkHot)
    {
        _overlayLinkHot = false;
        InvalidateRect(hwnd, nullptr, TRUE);
    }
}

void ViewerVLC::OnOverlayLButtonUp(HWND hwnd, POINT pt) noexcept
{
    if (PtInRect(&_overlayLinkRect, pt) == 0)
    {
        return;
    }

    const std::wstring url = GetOverlayLinkUrl();
    if (url.empty())
    {
        return;
    }

    static_cast<void>(ShellExecuteW(_hWnd ? _hWnd.get() : hwnd, L"open", url.c_str(), nullptr, nullptr, SW_SHOWNORMAL));
}

LRESULT ViewerVLC::OnOverlaySetCursor(HWND /*hwnd*/) noexcept
{
    if (! _overlayLinkHot)
    {
        return FALSE;
    }

    HCURSOR cursor = LoadCursorW(nullptr, IDC_HAND);
    if (! cursor)
    {
        cursor = LoadCursorW(nullptr, IDC_ARROW);
    }

    SetCursor(cursor);
    return TRUE;
}

std::wstring ViewerVLC::GetOverlayTitleText() const
{
    return LoadStringResource(g_hInstance, IDS_VIEWERVLC_MISSING_TITLE);
}

std::wstring ViewerVLC::GetOverlayBodyText() const
{
    std::wstring message = LoadStringResource(g_hInstance, IDS_VIEWERVLC_MISSING_BODY);
    if (! _overlayDetails.empty())
    {
        message = std::format(L"{}\r\n\r\n{}", message, _overlayDetails);
    }
    return message;
}

std::wstring ViewerVLC::GetOverlayLinkLabelText() const
{
    return LoadStringResource(g_hInstance, IDS_VIEWERVLC_MISSING_LINK_LABEL);
}

std::wstring ViewerVLC::GetOverlayLinkUrl() const
{
    return L"https://www.videolan.org/vlc/";
}

bool ViewerVLC::EnsureVlcLoaded(std::wstring& outError, bool enableAudioVisualization) noexcept
{
    outError.clear();

    std::filesystem::path installDir;
    if (! _config.vlcInstallPath.empty())
    {
        const std::filesystem::path configured = NormalizeVlcInstallPath(_config.vlcInstallPath);
        if (IsVlcInstallDir(configured))
        {
            installDir = configured;
        }
        else
        {
            outError = std::format(L"Configured VLC path is not a VLC installation folder: {}", configured.wstring());
        }
    }

    if (installDir.empty() && _config.autoDetectVlc)
    {
        const auto detected = AutoDetectVlcInstallDir();
        if (detected.has_value())
        {
            installDir = detected.value();
        }
    }

    if (installDir.empty())
    {
        if (outError.empty())
        {
            outError = L"VLC installation not found.";
        }
        return false;
    }

    const std::string pluginPathUtf8 = Utf8FromUtf16((installDir / L"plugins").wstring());

    std::vector<std::string> argStorage;
    argStorage.reserve(16);
    argStorage.emplace_back("--no-video-title-show");
    if (_config.quiet)
    {
        argStorage.emplace_back("--quiet");
    }

    if (! pluginPathUtf8.empty())
    {
        argStorage.push_back(std::format("--plugin-path={}", pluginPathUtf8));
    }

    if (_config.fileCachingMs > 0)
    {
        argStorage.push_back(std::format("--file-caching={}", _config.fileCachingMs));
    }

    if (_config.networkCachingMs > 0)
    {
        argStorage.push_back(std::format("--network-caching={}", _config.networkCachingMs));
    }

    if (! _config.avcodecHw.empty())
    {
        argStorage.push_back(std::format("--avcodec-hw={}", _config.avcodecHw));
    }

    if (! _config.videoOutput.empty())
    {
        argStorage.push_back(std::format("--vout={}", _config.videoOutput));
    }

    if (! _config.audioOutput.empty())
    {
        argStorage.push_back(std::format("--aout={}", _config.audioOutput));
    }

    if (enableAudioVisualization && ! _config.audioVisualization.empty() && _config.audioVisualization != "off")
    {
        argStorage.push_back(std::format("--audio-visual={}", _config.audioVisualization));
    }

    if (! _config.extraArgs.empty())
    {
        const auto extra = SplitVlcArgs(_config.extraArgs);
        for (const auto& a : extra)
        {
            if (! a.empty())
            {
                argStorage.push_back(a);
            }
        }
    }

    std::string desiredKey;
    desiredKey.reserve(argStorage.size() * 32);
    for (const auto& a : argStorage)
    {
        desiredKey.append(a);
        desiredKey.push_back('\n');
    }

    if (_vlc && _vlc->instance && _vlc->module && _vlc->installDir == installDir && _vlc->instanceArgsKey == desiredKey)
    {
        return true;
    }

    StopPlayback();
    _vlc.reset();

    auto state             = std::make_unique<VlcState>();
    state->installDir      = installDir;
    state->instanceArgsKey = desiredKey;

    const std::filesystem::path dllPath = installDir / L"libvlc.dll";

    const DWORD prevNeeded = GetDllDirectoryW(0, nullptr);
    if (prevNeeded > 0 && prevNeeded < 32768)
    {
        std::wstring prev(static_cast<size_t>(prevNeeded), L'\0');
        const DWORD prevWritten = GetDllDirectoryW(prevNeeded, prev.data());
        if (prevWritten > 0 && prevWritten < prevNeeded)
        {
            prev.resize(static_cast<size_t>(prevWritten));
            state->previousDllDirectory = std::move(prev);
        }
    }
    state->dllDirectoryWasSet = SetDllDirectoryW(installDir.c_str()) != 0;

    HMODULE module = LoadLibraryW(dllPath.c_str());
    if (! module)
    {
        const DWORD lastError = GetLastError();
        outError              = std::format(L"Failed to load '{}' (Win32: {}).", dllPath.wstring(), lastError);
        return false;
    }

    state->module.reset(module);

    bool ok = true;
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_new", state->libvlc_new);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_release", state->libvlc_release);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_media_new_path", state->libvlc_media_new_path);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_media_release", state->libvlc_media_release);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_media_player_new_from_media", state->libvlc_media_player_new_from_media);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_media_player_release", state->libvlc_media_player_release);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_media_player_set_hwnd", state->libvlc_media_player_set_hwnd);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_media_player_play", state->libvlc_media_player_play);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_media_player_pause", state->libvlc_media_player_pause);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_media_player_stop", state->libvlc_media_player_stop);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_media_player_is_playing", state->libvlc_media_player_is_playing);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_media_player_get_time", state->libvlc_media_player_get_time);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_media_player_set_time", state->libvlc_media_player_set_time);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_media_player_get_length", state->libvlc_media_player_get_length);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_audio_set_volume", state->libvlc_audio_set_volume);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_audio_get_volume", state->libvlc_audio_get_volume);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_video_take_snapshot", state->libvlc_video_take_snapshot);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_media_player_set_rate", state->libvlc_media_player_set_rate);
    ok      = ok && TryLoadProc(state->module.get(), "libvlc_media_player_get_rate", state->libvlc_media_player_get_rate);

    if (! ok)
    {
        outError = L"Failed to resolve libVLC exports from libvlc.dll.";
        return false;
    }

    state->instance.get_deleter().release = state->libvlc_release;
    state->player.get_deleter().release   = state->libvlc_media_player_release;

    std::vector<const char*> argv;
    argv.reserve(argStorage.size());
    for (auto& a : argStorage)
    {
        argv.push_back(a.c_str());
    }

    libvlc_instance_t* inst = state->libvlc_new(static_cast<int>(argv.size()), argv.data());
    if (! inst)
    {
        outError = L"libvlc_new failed.";
        return false;
    }

    state->instance.reset(inst);

    _vlc = std::move(state);
    return true;
}

bool ViewerVLC::StartPlayback(const std::filesystem::path& path) noexcept
{
    StopPlayback();

    if (path.empty())
    {
        SetMissingUiVisible(true, L"File path is empty.");
        return false;
    }

    std::error_code ec;
    if (! std::filesystem::exists(path, ec) || ! std::filesystem::is_regular_file(path, ec))
    {
        SetMissingUiVisible(true, L"This file is not available as a local file path.");
        return false;
    }

    _hudRate = std::clamp(static_cast<float>(_config.defaultPlaybackRatePercent) / 100.0f, 0.25f, 4.0f);

    std::wstring error;
    _isAudioFile                        = IsAudioExtension(path.extension().wstring());
    const bool enableAudioVisualization = _isAudioFile && ! _config.audioVisualization.empty() && _config.audioVisualization != "off";

    if (! EnsureVlcLoaded(error, enableAudioVisualization))
    {
        SetMissingUiVisible(true, error);
        return false;
    }

    if (! _vlc || ! _vlc->instance || ! _vlc->libvlc_media_new_path || ! _vlc->libvlc_media_player_new_from_media || ! _vlc->libvlc_media_release ||
        ! _vlc->libvlc_media_player_release)
    {
        SetMissingUiVisible(true, L"libVLC is not available.");
        return false;
    }

    const std::string pathUtf8 = Utf8FromUtf16(path.wstring());
    if (pathUtf8.empty())
    {
        SetMissingUiVisible(true, L"Failed to convert the file path to UTF-8.");
        return false;
    }

    std::unique_ptr<libvlc_media_t, VlcState::MediaDeleter> media(nullptr, {_vlc->libvlc_media_release});
    media.reset(_vlc->libvlc_media_new_path(_vlc->instance.get(), pathUtf8.c_str()));
    if (! media)
    {
        SetMissingUiVisible(true, L"libvlc_media_new_path failed.");
        return false;
    }

    std::unique_ptr<libvlc_media_player_t, VlcState::PlayerDeleter> player(nullptr, {_vlc->libvlc_media_player_release});
    player.reset(_vlc->libvlc_media_player_new_from_media(media.get()));
    if (! player)
    {
        SetMissingUiVisible(true, L"libvlc_media_player_new_from_media failed.");
        return false;
    }

    if (_vlc->libvlc_media_player_set_hwnd && _hVideo)
    {
        _vlc->libvlc_media_player_set_hwnd(player.get(), _hVideo.get());
    }

    if (_vlc->libvlc_audio_set_volume)
    {
        _vlc->libvlc_audio_set_volume(player.get(), std::clamp(_hudVolumeValue, 0, 100));
    }

    if (_vlc->libvlc_media_player_set_rate)
    {
        _vlc->libvlc_media_player_set_rate(player.get(), std::clamp(_hudRate, 0.25f, 4.0f));
    }

    _vlc->player = std::move(player);

    SetMissingUiVisible(false, {});

    if (_vlc->libvlc_media_player_play)
    {
        const int rc = _vlc->libvlc_media_player_play(_vlc->player.get());
        if (rc != 0)
        {
            SetMissingUiVisible(true, std::format(L"libvlc_media_player_play failed (code {}).", rc));
            _vlc->player.reset();
            return false;
        }
    }

    if (_hWnd)
    {
        _uiTimerId = SetTimer(_hWnd.get(), kUiTimerId, kUiTimerIntervalMs, nullptr);
    }

    _hudLastActivityTick = GetTickCount64();

    UpdatePlaybackUi();
    return true;
}

void ViewerVLC::StopPlayback() noexcept
{
    if (_hWnd && _uiTimerId != 0)
    {
        KillTimer(_hWnd.get(), _uiTimerId);
        _uiTimerId = 0;
    }

    _hudSeekDragging    = false;
    _hudVolumeDragging  = false;
    _hudPressed         = HudPart::None;
    _hudHot             = HudPart::None;
    _hudDragTimeMs      = 0;
    _seekDragWasPlaying = false;
    ClearSeekPreview();

    if (_vlc && _vlc->player)
    {
        if (_vlc->libvlc_media_player_stop)
        {
            _vlc->libvlc_media_player_stop(_vlc->player.get());
        }

        _vlc->player.reset();
    }

    UpdatePlaybackUi();
}

void ViewerVLC::TakeSnapshot() noexcept
{
    if (! _vlc || ! _vlc->player || ! _vlc->libvlc_video_take_snapshot || ! _hWnd)
    {
        return;
    }

    std::array<wchar_t, 2048> fileBuffer{};
    fileBuffer[0] = L'\0';

    const std::wstring filter = LoadStringResource(g_hInstance, IDS_VIEWERVLC_FILEDLG_FILTER_PNG);
    const std::wstring title  = LoadStringResource(g_hInstance, IDS_VIEWERVLC_FILEDLG_TITLE_SNAPSHOT);

    OPENFILENAMEW ofn{};
    ofn.lStructSize  = sizeof(ofn);
    ofn.hwndOwner    = _hWnd.get();
    ofn.lpstrFile    = fileBuffer.data();
    ofn.nMaxFile     = static_cast<DWORD>(fileBuffer.size());
    ofn.lpstrFilter  = filter.c_str();
    ofn.nFilterIndex = 1;
    ofn.lpstrDefExt  = L"png";
    ofn.lpstrTitle   = title.empty() ? nullptr : title.c_str();
    ofn.Flags        = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_EXPLORER | OFN_NOCHANGEDIR | OFN_HIDEREADONLY;

    if (! GetSaveFileNameW(&ofn))
    {
        return;
    }

    const std::filesystem::path outPath(fileBuffer.data());
    const std::string outUtf8 = Utf8FromUtf16(outPath.wstring());
    if (outUtf8.empty())
    {
        return;
    }

    const int rc = _vlc->libvlc_video_take_snapshot(_vlc->player.get(), 0, outUtf8.c_str(), 0, 0);
    if (rc == 0)
    {
        return;
    }

    if (! _hostAlerts)
    {
        return;
    }

    const std::wstring message = std::format(L"Snapshot failed (code {}).", rc);

    HostAlertRequest request{};
    request.version      = 1;
    request.sizeBytes    = sizeof(request);
    request.scope        = HOST_ALERT_SCOPE_WINDOW;
    request.modality     = HOST_ALERT_MODELESS;
    request.severity     = HOST_ALERT_ERROR;
    request.targetWindow = _hWnd.get();
    request.title        = nullptr;
    request.message      = message.c_str();
    request.closable     = TRUE;

    static_cast<void>(_hostAlerts->ShowAlert(&request, _hWnd.get()));
}

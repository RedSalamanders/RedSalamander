#pragma once

#include <optional>

#include "Helpers.h"
#include "SettingsStore.h"

namespace Common::WindowBackdrop
{
enum class Kind : int
{
    None    = 1,
    Mica    = 2,
    Acrylic = 3,
    MicaAlt = 4,
};

enum class Target : uint8_t
{
    Primary,
    Tool,
};

[[nodiscard]] constexpr Kind Resolve(Common::Settings::WindowBackdropMode mode, Target target, bool highContrast) noexcept
{
    if (highContrast)
    {
        return Kind::None;
    }

    switch (mode)
    {
        case Common::Settings::WindowBackdropMode::None: return Kind::None;
        case Common::Settings::WindowBackdropMode::Mica: return Kind::Mica;
        case Common::Settings::WindowBackdropMode::MicaAlt: return Kind::MicaAlt;
        case Common::Settings::WindowBackdropMode::Acrylic: return Kind::Acrylic;
        case Common::Settings::WindowBackdropMode::Default:
        default: return target == Target::Primary ? Kind::Mica : Kind::MicaAlt;
    }
}

[[nodiscard]] constexpr int ToDwmSystemBackdropValue(Kind kind) noexcept
{
    return static_cast<int>(kind);
}

inline HRESULT SetWindowBackdropKind(HWND hwnd, Kind kind) noexcept
{
    if (! hwnd)
    {
        return E_INVALIDARG;
    }

    using DwmSetWindowAttributeFunc          = HRESULT(WINAPI*)(HWND, DWORD, LPCVOID, DWORD);
    static DwmSetWindowAttributeFunc setAttr = []() noexcept -> DwmSetWindowAttributeFunc
    {
        HMODULE dwm = LoadLibraryW(L"dwmapi.dll");
        if (! dwm)
        {
            return nullptr;
        }
#pragma warning(push)
#pragma warning(disable : 4191)
        return reinterpret_cast<DwmSetWindowAttributeFunc>(GetProcAddress(dwm, "DwmSetWindowAttribute"));
#pragma warning(pop)
    }();

    if (! setAttr)
    {
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }

    constexpr DWORD kDwmwaSystemBackdropType = 38u;
    const int value                          = ToDwmSystemBackdropValue(kind);
    return setAttr(hwnd, kDwmwaSystemBackdropType, &value, sizeof(value));
}

inline void ApplyWindowBackdropKind(HWND hwnd, Kind kind) noexcept
{
    const HRESULT hr = SetWindowBackdropKind(hwnd, kind);
    if (FAILED(hr) && kind != Kind::None)
    {
        static_cast<void>(SetWindowBackdropKind(hwnd, Kind::None));
    }
}

inline void ApplyResolvedWindowBackdrop(HWND hwnd, Common::Settings::WindowBackdropMode mode, Target target, bool highContrast) noexcept
{
    ApplyWindowBackdropKind(hwnd, Resolve(mode, target, highContrast));
}

inline std::optional<Kind> TryGetAppliedWindowBackdropKind(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return std::nullopt;
    }

    using DwmGetWindowAttributeFunc          = HRESULT(WINAPI*)(HWND, DWORD, PVOID, DWORD);
    static DwmGetWindowAttributeFunc getAttr = []() noexcept -> DwmGetWindowAttributeFunc
    {
        HMODULE dwm = LoadLibraryW(L"dwmapi.dll");
        if (! dwm)
        {
            return nullptr;
        }
#pragma warning(push)
#pragma warning(disable : 4191)
        return reinterpret_cast<DwmGetWindowAttributeFunc>(GetProcAddress(dwm, "DwmGetWindowAttribute"));
#pragma warning(pop)
    }();

    if (! getAttr)
    {
        return std::nullopt;
    }

    constexpr DWORD kDwmwaSystemBackdropType = 38u;
    int value                                = 0;
    if (FAILED(getAttr(hwnd, kDwmwaSystemBackdropType, &value, sizeof(value))))
    {
        return std::nullopt;
    }

    switch (value)
    {
        case 1: return Kind::None;
        case 2: return Kind::Mica;
        case 3: return Kind::Acrylic;
        case 4: return Kind::MicaAlt;
        default: return std::nullopt;
    }
}
} // namespace Common::WindowBackdrop

#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#include <array>
#include <cstdint>
#include <optional>
#include <string>

#include "AppTheme.h"

#include <Windows.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#pragma warning(pop)

class ShortcutManager;

class FunctionBar
{
public:
    FunctionBar();
    ~FunctionBar();

    FunctionBar(const FunctionBar&)            = delete;
    FunctionBar& operator=(const FunctionBar&) = delete;
    FunctionBar(FunctionBar&&)                 = delete;
    FunctionBar& operator=(FunctionBar&&)      = delete;

    HWND Create(HWND parent, int x, int y, int width, int height);
    void Destroy() noexcept;

    [[nodiscard]] HWND GetHwnd() const noexcept
    {
        return _hWnd.get();
    }

    void SetTheme(const AppTheme& theme);
    void SetShortcutManager(const ShortcutManager* shortcuts) noexcept;
    void SetDpi(UINT dpi) noexcept;
    void SetModifiers(uint32_t modifiers) noexcept;
    void SetPressedFunctionKey(std::optional<uint32_t> vk) noexcept;

private:
    static ATOM RegisterWndClass(HINSTANCE instance);
    static constexpr PCWSTR kClassName = L"RedSalamander.FunctionBar";

    static LRESULT CALLBACK WndProcThunk(HWND hWindow, UINT msg, WPARAM wp, LPARAM lp);
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);

    void OnCreate(HWND hwnd) noexcept;
    void OnDestroy() noexcept;
    void OnPaint() noexcept;
    void OnSize(UINT width, UINT height) noexcept;
    void OnMouseMove(POINT pt) noexcept;
    void OnMouseLeave() noexcept;
    void OnLButtonUp(POINT pt) noexcept;

    void RecomputeLabels();
    void RecomputeModifierText();
    [[nodiscard]] int PxFromDip(int dip) const noexcept;
    [[nodiscard]] std::optional<uint32_t> HitTestFunctionKey(POINT pt) const noexcept;
    void EnsureKeyFont() noexcept;
    void EnsureTextFont() noexcept;

private:
    wil::unique_hwnd _hWnd;
    HINSTANCE _hInstance = nullptr;
    UINT _dpi            = USER_DEFAULT_SCREEN_DPI;
    SIZE _clientSize{};

    AppTheme _theme;
    const ShortcutManager* _shortcuts = nullptr;
    uint32_t _modifiers               = 0;
    std::optional<uint32_t> _pressedKey;
    std::optional<uint32_t> _hoveredKey;
    bool _trackingMouseLeave = false;

    std::array<std::wstring, 12> _labels{};
    std::wstring _modifierText;

    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> _backgroundBrush;
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> _pressedBrush;
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> _hoverBrush;
    wil::unique_any<HPEN, decltype(&::DeleteObject), ::DeleteObject> _glyphPen;
    wil::unique_any<HPEN, decltype(&::DeleteObject), ::DeleteObject> _separatorPen;
    wil::unique_any<HFONT, decltype(&::DeleteObject), ::DeleteObject> _keyFont;
    wil::unique_any<HFONT, decltype(&::DeleteObject), ::DeleteObject> _textFont;
};

#pragma once

#include "AppTheme.h"

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

namespace ThemedInputFrames
{
struct FrameStyle
{
    const AppTheme* theme                 = nullptr;
    HBRUSH backdropBrush                  = nullptr;
    COLORREF inputBackgroundColor         = RGB(255, 255, 255);
    COLORREF inputFocusedBackgroundColor  = RGB(255, 255, 255);
    COLORREF inputDisabledBackgroundColor = RGB(255, 255, 255);
};

void InstallFrame(HWND frame, HWND input, FrameStyle* style) noexcept;
void InstallControl(HWND input, HWND frame = nullptr) noexcept;

void InvalidateComboBox(HWND combo) noexcept;

LRESULT HandleInputControlMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, HWND frame, bool& handled) noexcept;
LRESULT HandleInputFrameMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, FrameStyle* style, bool& handled) noexcept;

} // namespace ThemedInputFrames

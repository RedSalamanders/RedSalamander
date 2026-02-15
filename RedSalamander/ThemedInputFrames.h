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

void InvalidateComboBox(HWND combo) noexcept;

LRESULT CALLBACK InputControlSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData) noexcept;
LRESULT CALLBACK InputFrameSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData) noexcept;
} // namespace ThemedInputFrames

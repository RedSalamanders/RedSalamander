Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Describe 'HWND render-target resource contracts' {
    BeforeAll {
        $sharedSource = Get-Content -LiteralPath (Join-Path $repoRoot 'Common\HwndRenderTargetResources.h') -Raw
        $functionBarSource = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\FunctionBar.cpp') -Raw
        $statusBarSource = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\FolderWindow.StatusBar.cpp') -Raw
    }

    It 'owns the common factory target and brush lifecycle in one narrow helper' {
        $sharedSource | Should Match 'struct\s+HwndRenderTargetResources'
        $sharedSource | Should Match 'wil::com_ptr<ID2D1Factory>\s+d2dFactory'
        $sharedSource | Should Match 'wil::com_ptr<ID2D1HwndRenderTarget>\s+target'
        $sharedSource | Should Match 'wil::com_ptr<ID2D1SolidColorBrush>\s+solidBrush'
        $sharedSource | Should Match 'void\s+ResetTarget\(\)\s+noexcept\s*\{\s*solidBrush\.reset\(\);\s*target\.reset\(\);'
        $sharedSource | Should Match '\[\[nodiscard\]\]\s+bool\s+EnsureD2dFactory\(\)\s+noexcept'
        $sharedSource | Should Match '\[\[nodiscard\]\]\s+bool\s+EnsureTarget\(HWND hwnd, UINT dpi\)\s+noexcept'
        ([regex]::Matches($sharedSource, 'CreateHwndRenderTarget\(')).Count | Should Be 1
    }

    It 'keeps FunctionBar and StatusBar drawing policy local while delegating exact lifecycle work' {
        foreach ($source in @($functionBarSource, $statusBarSource)) {
            $source | Should Match '#include "HwndRenderTargetResources\.h"'
            $source | Should Match 'RenderResources\s+final\s*:\s*Common::Rendering::HwndRenderTargetResources'
            $source | Should Match 'resources\.EnsureD2dFactory\(\)'
            $source | Should Match 'resources\.EnsureTarget\(hwnd, resources\.dpi\)'
            $source | Should Match 'resources\.ResetTarget\(\)'
            $source | Should Not Match 'CreateHwndRenderTarget\('
            $source | Should Not Match 'wil::com_ptr<ID2D1HwndRenderTarget>\s+target'
            $source | Should Not Match 'wil::com_ptr<ID2D1SolidColorBrush>\s+solidBrush'
        }
    }

    It 'retains surface-specific typography invalidation and HWND teardown ownership' {
        $functionBarSource | Should Match 'void\s+ResetFunctionBarTypography[\s\S]*?ResetFunctionBarTarget\(resources\);'
        $statusBarSource | Should Match 'void\s+ResetStatusBarTypography[\s\S]*?ResetStatusBarTarget\(resources\);'
        $functionBarSource | Should Match 'void\s+FunctionBar::OnDestroy\(\)\s+noexcept\s*\{\s*DestroyFunctionBarRenderResources'
        $statusBarSource | Should Match 'StatusBarOnNcDestroy[\s\S]*?DestroyStatusBarRenderResources\(hwnd\)'
        $functionBarSource | Should Match 'Debug::Perf::Scope\s+paintPerf\(L"render\.function_bar\.paint_us"\)'
        $statusBarSource | Should Match 'Debug::Perf::Scope\s+paintPerf\(L"render\.status_bar\.paint_us"\)'
    }
}

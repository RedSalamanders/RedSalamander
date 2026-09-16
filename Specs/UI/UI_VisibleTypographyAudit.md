# Visible Typography Contract

Status: current normative audit contract
Last reviewed: 2026-08-04

## Product contract

- App-owned visible text uses the typography services in the public `<DxUi/Typography.h>` and DirectWrite.
- Windows 11 UI families are `Segoe UI Variable Small` for caption/header-scale text, `Segoe UI Variable Text` for body/control text, and `Segoe UI Variable Display` for large display text.
- `Segoe Fluent Icons`, `Segoe UI Emoji`, and `Consolas` are the only family exceptions for icon glyphs, emoji, and monospace content respectively.
- Product code must not introduce visible GDI text, caller-owned UI `HFONT` state, or native-font propagation as a fallback for a DxUi surface.
- Shared native interop that remains necessary belongs in the reviewed shared bridge named by `UI_VisibleNativeAudit.md`; consumers must not recreate it locally.

## Current exception inventory

`Tools/Get-VisibleTypographyAudit.ps1` currently reports one non-product hit:

| Path | Exception | Rationale and owner |
|---|---|---|
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.cpp` | `DirectedSelfTestInputWarning` creates a large warning font | Test-only, `ENABLE_TESTS`-guarded desktop-input safety UI. The canonical shared test helper and restoration requirements are owned by `Core_SharedHelpers.md` and `Testing_SelfTests.md`. The font is held by WIL RAII. |

There are no current product-surface typography exceptions. Any new audit result outside the row above is drift and must be resolved in code or explicitly reviewed here before merge.

## Native text input

DxUi `TextField` and editable `ComboBox` input use the host-HWND retained native text-input session. NavigationView address editing uses that DxUi-backed host. Hidden Edit/RichEdit bridge windows and bridge-owned font state are not part of the current architecture.

## Verification

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\Tools\Get-VisibleTypographyAudit.ps1
```

The audit scans production and test sources. Changes to typography, visible GDI/Win32 rendering, native text input, or the audit allowlist must also run the relevant DxUi/Commands/viewer selftests and update the authoritative owning spec when behavior changes.

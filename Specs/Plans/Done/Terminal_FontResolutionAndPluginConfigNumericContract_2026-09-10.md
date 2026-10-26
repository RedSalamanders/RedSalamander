# I20 Terminal font resolution and the plugin-configuration numeric contract

**Status:** complete — recorded pass cited 2026-10-17; this file moved to `Specs/Plans/Done/`
**Opened:** 2026-09-10
**Closed:** 2026-10-17

## Problem

The embedded terminal rendered Nerd Font glyphs (powerline separators, folder icons) as `.notdef`
boxes even though the persisted configuration correctly held `"fontFamily": "Cascadia Code NF"`.
Measured on the reporting machine through DirectWrite: `Cascadia Code NF` resolves to
`C:\Windows\Fonts\CascadiaCodeNF.ttf` and maps U+E0B0, U+E0B4, U+E5FF, U+F07B and U+F0F1D;
`Cascadia Mono` resolves to `CascadiaMono.ttf` and maps none of them. Both families are truly
monospaced — every probed glyph shares the `M` advance of 0.5859 em — so cell geometry was never
involved.

## Root cause

`Terminal::SetConfiguration` guarded `fontSizeDip` with `yyjson_is_num` (true for int/uint/real)
and then read it with `yyjson_get_real`, which yyjson documents as returning `0.0` unless the
value's storage subtype is real. A JSON `14` parses as uint, so the read produced `0.0`, failed the
`>= 8.0` range check, and returned `E_INVALIDARG`.

Because `SetConfiguration` is transactional, that single member discarded the entire candidate —
including `fontFamily` — so `_config` kept its compiled default `L"Cascadia Mono"`, which reached
`CreateTextFormat`. Nerd Font glyphs live in the Unicode private use areas, which deliberately have
no system font fallback, so the missing glyphs became `.notdef` boxes rather than substituting.

Three things made it permanent and invisible:

1. **The writer can only emit an integer.** The Terminal's own schema declares `fontSizeDip` as a
   `"value"` field with integer default/min/max, and both the Preferences page and the shared
   `Common::PluginConfiguration` codec serialize `value` fields as `int64_t`. The plugin therefore
   refused its own schema defaults. It appeared to work only because `GetConfiguration` emits
   `14.0`; the first Preferences round-trip normalized that to `14` and it never recovered.
2. **The Preferences page never asked the plugin.** `CommitEditor` wrote straight into
   `workingSettings`, unlike every other writer, so a configuration the plugin would reject could
   be persisted.
3. **Every caller voided the HRESULT.** All four `ApplyConfigurationFromSettings` call sites
   discarded the result, so the rejection produced no log line.

No test caught it: `RequiresTransactionalConfigurationProof` listed only the six file-system
providers, and nothing drove a plugin's own schema defaults back through `SetConfiguration`. The
one existing terminal-config selftest fixture used `{"defaultShell":"cmd"}` and never touched
`fontSizeDip`.

## Changes

**Canonical numeric accessor.** `Common/YyjsonHelpers.h` had `GetInt64Member` / `GetUInt64Member`
but no double accessor, which is why the read was hand-rolled. Added `Common::Json::GetDoubleMember`
(and `ParseRealString`), accepting sint/uint/real uniformly and reporting non-finite values or
magnitudes that cannot round-trip the 53-bit mantissa as `OutOfRange`. The other six
`yyjson_get_real` call sites in the repository were audited and are correct — each is guarded by
`yyjson_is_real` with explicit integer branches — and were left untouched.

**Terminal configuration.** `Terminal::SetConfiguration` now parses through
`Common::Json::ParseObjectDocument` and reads each member with the shared accessors, matching
`Plugins/FileSystemS3/FileSystemS3.Configuration.cpp`. `ReadBoolean` stays as a documented local
variant because the schema's `option` fields store `"0"`/`"1"` strings that `GetBoolMember` rejects.
Malformed JSON and non-object roots now return `HRESULT_FROM_WIN32(ERROR_INVALID_DATA)` as the
plugin contract requires, and unknown members are copied through in their original order — both
were pre-existing non-conformances that adding the Terminal to the transactional proof exposed.

**Font resolution and the user-visible notice.** `Terminal::resolveFontFamily` walks a
monospace-only chain (configured family → `Cascadia Mono` → `Consolas`) using
`Typography::IsFontFamilyAvailable`. The shared `Typography::CreateTextFormat` is deliberately not
used: its fallback resolves unknown families to the proportional `Segoe UI`, which cannot hold a
character grid. Glyph coverage is probed once per resolved family over U+E0B0/U+E0B4/U+F07B via
`IDWriteFontFace::GetGlyphIndices`, and only reported once the screen actually paints a private-use
codepoint. `drawFontNotice` paints a non-blocking, dismissible single-line banner following the
visual conventions of `drawConfirmation` without its dim or modal gate; the text is also appended
to the accessibility snapshot. `setDiagnostic` was deliberately avoided because a non-empty
diagnostic replaces the whole grid.

**Font-collection freshness.** `QueryFontFamilyAvailable` now passes `checkForUpdates = TRUE`, and
`Typography::InvalidateFontFamilyAvailability` drops the memoized answers for one factory. The
Terminal calls it when the configured family changes, so a font installed while RedSalamander is
running is picked up without a restart.

**Validated writes and diagnostics.** `FileSystemPluginManager::ValidateConfiguration` and
`ViewerPluginManager::ValidateConfiguration` check a candidate on a throwaway instance — no mounted
provider or open pane is disturbed — and Preferences' `CommitEditor` refuses to persist a rejected
candidate, reporting it through the existing field-error surface
(`IDS_PREFS_PLUGINS_DETAILS_CONFIG_REJECTED`). Both `ApplyConfigurationFromSettings`
implementations now log a `Debug::Warning` naming the plugin id and `HRESULT` instead of voiding it.

**Coverage.** `Tests/PluginContractTests` gained `TestSchemaDefaultsRoundTrip`, applied to every
plugin exposing `IInformations`: schema defaults materialized through the exported
`Common::PluginConfiguration::MakeDefaultValue` and `SerializeConfiguration` must be accepted with
`S_OK`, and the resulting `GetConfiguration` output must still satisfy the schema and be
re-acceptable. The Terminal was added to the existing transactional proofs.
`settings_terminal_plugin_roundtrip` covers the host plumbing, and
`TestDxUiTypographyFontAvailabilityRescansAfterInvalidation` covers the typography cache.
`Tools/Tests/TestHarnessSourceContracts.Tests.ps1` bans a bare `yyjson_get_real` in
`Plugins/Terminal/Terminal.cpp` and pins the new assertion names.

## Performance

No new per-frame work. The protected scenario is terminal device-resource recreation and
configuration apply. Family availability and glyph coverage resolve at most once per family per
factory behind the existing `IsFontFamilyAvailable` cache, which already emits
`dxui.typography.family_cache_miss_count`. The per-cell hot path gains one boolean test
(`! _iconGlyphRequested`) that latches on the first private-use cell, so an all-ASCII screen never
runs the scan. The banner draws only while visible.

## Closeout

Durable requirements were merged into the authoritative specs before this plan moves to
`Specs/Plans/Done/`:

- `Specs/Plugins/Plugins_VirtualFileSystem.md` — numeric acceptance (a JSON integer is a valid
  value for any numeric member), the shared-accessor requirement, and the rule that a host must not
  persist a configuration the plugin has not accepted.
- `Specs/Terminal/Terminal_EmbeddedPlugin.md` — the `fontSizeDip` integer-or-real validation, the
  `ERROR_INVALID_DATA` and unknown-member behaviour, and the font resolution / notice contract.
- `Specs/Core/Core_SharedHelpers.md` — `GetDoubleMember` and the `Common/PluginConfiguration.h`
  schema-defaults entry.
- `Specs/Testing/Testing_TestCoverage.md` — the schema-defaults conformance requirement and the
  terminal font coverage requirement.

## Verification citation (2026-10-17)

The closeout rule above is the durable-spec merge. Those four specs already contain the
numeric, font, shared-helper, and coverage rules. This citation is the recorded pass that
was missing from the plan; it does not add a new requirement.

Release Full/Fresh `20261001T115346Z-104056-8e2cfb6a0b8049218c622dd915a2697e` passed
**2,211 / 0 failed / 46 skipped** on snapshot
`f2e0a18a04fbc8a95368adc0c946bafb5ebccdf68d81b4a476d0fefbd323593c`, git head
`ccebd5af68eb45f2e9deed64a7063780e71eb1ce`. That head contains the fix commit
`6dbcedeb7` (2026-09-10). The packet is
`Specs/TestRuns/4cb089111a23/Validation/2026-10-01_i26-v65-final/`; case text is in
ZIP member `full/commands-native-results.json` and `full/runner.log` inside
`compressed/run-text-evidence.zip`.

- `settings_terminal_plugin_roundtrip` passed in 2 ms. It is registered in
  `RedSalamander/SelfTest/Commands/Commands.Tests.PluginConfig.cpp`
  (`RunPluginConfigCommandsSelfTestCases`).
- `TestSchemaDefaultsRoundTrip` is a helper in
  `Tests/PluginContractTests/PluginContract.Tests.Runner.cpp`, called for every plugin
  whose schema is present. The same run logged `PluginContractTests passed` and
  `[       OK ] Plugins\Terminal.dll: SetConfiguration accepts its own schema defaults for pluginId=builtin/terminal`,
  plus the matching shared-codec serialization line. No schema-defaults assertion failed.
- The September 8 Full `20260908T200723Z-69288-1db8a2860de7474ebbc2f9164aa99216` does not
  contain `settings_terminal_plugin_roundtrip`. That case landed with `6dbcedeb7`, after
  that run.
- `TestDxUiTypographyFontAvailabilityRescansAfterInvalidation` is not in this
  repository's test sources, and it was not found in the `Specs/TestRuns` archives
  searched for this closeout. It is not the Done gate. The product Full above is the
  recorded pass of the plugin-contract and Commands cases.
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` still bans a bare `yyjson_get_real`
  in `Plugins/Terminal/Terminal.cpp` and pins `GetDoubleMember`, the schema-defaults
  assertion text, `resolveFontFamily`, and `drawFontNotice`. This closeout did not
  re-run that Pester file; the Full's promoted `tooling.pester` entry is not cited as
  that individual test.

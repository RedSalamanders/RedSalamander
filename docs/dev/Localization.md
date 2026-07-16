# Localization

RedSalamander keeps every user-facing string and every static menu, dialog, and accelerator in Windows `.rc` resources, ships English embedded in each binary, and overlays translations through per-culture satellite DLLs. This page is the developer reference for how that system is laid out, the runtime lookup and fallback chain, the placeholder rules translators must follow, and the day-to-day workflows for adding a string and adding a culture.

For the public/onboarding overview see the [Developer Guide](../DeveloperGuide.md). The normative behavior lives in `Specs/Core/Core_Localization.md`, and the hands-on patterns are in `.github/skills/localization/SKILL.md`.

## Where strings live (.rc ownership)

Each executable and plugin owns the English resources it uses, embedded directly in that module. Nothing user-facing is hardcoded in C++.

| Owner module | Embedded `.rc` |
| --- | --- |
| RedSalamander (main app) | `RedSalamander/RedSalamander.rc` |
| RedSalamanderMonitor | `RedSalamanderMonitor/RedSalamanderMonitor.rc` |
| RedConfigure | `RedConfigure/RedConfigure.rc` |
| Each plugin DLL | `Plugins/<PluginName>/<PluginName>.rc` |

The main app `.rc` carries `LANGUAGE LANG_ENGLISH, SUBLANG_ENGLISH_US`, menus (for example `IDR_FOLDERVIEW_CONTEXT`), dialogs (for example `IDD_PREFERENCES`), an accelerators table, and many `STRINGTABLE` blocks. Resource IDs are defined in the owner's `Resource.h` (the main app uses `RedSalamander/Resource.h`).

Rules of ownership:

- A string lives in exactly one owner's embedded `.rc`. Plugins keep their own strings; the main app does not host plugin strings.
- Satellite translations may translate any owner string except documented language-neutral IDs (see below).
- Static menu and dialog structure stays in `.rc`. Runtime code may only extend genuinely dynamic sections (the theme list, settings-defined custom themes, and per-user history entries).

## Satellite DLL naming and layout

English stays embedded in the owner module. Translations are resource-only DLLs, one per owner per culture, kept next to the owner under `Lang/<culture>/`.

```text
RedSalamander/
  RedSalamander.rc                         # embedded English (owner)
  Resource.h
  Lang/
    fr-FR/
      RedSalamander-fr-FR.rc               # French translation
      RedSalamander-fr-FR.vcxproj          # resource-only DLL project
    ja-JP/
      RedSalamander-ja-JP.rc
      RedSalamander-ja-JP.vcxproj
    cs-CZ/ ...
    sk-SK/ ...
Plugins/ViewerText/
  ViewerText.rc
  Lang/fr-FR/ViewerText-fr-FR.rc + .vcxproj
```

Naming and build conventions:

- A satellite DLL is named `<OwnerStem>-<culture>.dll`, where `<OwnerStem>` matches the owning executable or plugin DLL stem, for example `RedSalamander-fr-FR.dll`, `RedSalamanderMonitor-fr-FR.dll`, or `ViewerText-fr-FR.dll`.
- Each culture folder holds a generated `.rc` plus a `.vcxproj`. The project is a resource-only `DynamicLibrary` with no executable entry point; it marks itself with the localization properties consumed by the build:

  ```xml
  <PropertyGroup Label="Localization">
    <IsLanguageResourceProject>true</IsLanguageResourceProject>
    <ResourceOwnerName>RedSalamander</ResourceOwnerName>
    <ResourceCulture>fr-FR</ResourceCulture>
  </PropertyGroup>
  ```

- Satellite projects output to `.build\<Platform>\<Configuration>\Lang\`. Normal executable and plugin binaries land where they always did; only the satellites go into `Lang\`. `build.ps1` validates that language resource projects land in the `Lang\` output folder.
- A satellite `.rc` declares its own `LANGUAGE` (for example `LANGUAGE LANG_FRENCH, SUBLANG_FRENCH`), uses `#pragma code_page(65001)` for UTF-8 text, and includes the owner's `Resource.h` so the IDs match the embedded source.

The cultures currently shipped across owners are `cs-CZ`, `fr-FR`, `ja-JP`, and `sk-SK`.

## Runtime lookup and fallback chain

All localized resource lookup goes through the `Localization` namespace declared in [Common/LocalizationManager.h](../../Common/LocalizationManager.h). App and Common code normally call the thin wrappers in `Common/Helpers.h` (`LoadStringResource`, `FormatStringResource`, `MessageBoxResource`), which route through the manager.

### Owner registration

Before any localized lookup, each owner registers its embedded module so the manager can find both the embedded English and the matching satellites:

```cpp
HRESULT RegisterResourceOwner(std::wstring_view ownerName, HINSTANCE embeddedInstance) noexcept;
void    UnregisterResourceOwner(HINSTANCE embeddedInstance) noexcept;
```

- `RedSalamander` and `RedSalamanderMonitor` register during startup, after settings load and before UI resource lookup.
- Plugins register immediately after the DLL loads and unregister before the module unloads. Plugin code passes the plugin DLL handle (`g_hInstance`), never `nullptr`.

### Selecting the language

The persisted setting is `ui.language` in the settings file: either `system` (default) or a BCP-style culture tag matching `^[a-zA-Z]{2,3}(-[a-zA-Z0-9]{2,8})*$`, for example `fr` or `fr-FR`. The app maps that string into a `LanguagePreference` and calls `ApplyLanguagePreference`:

```cpp
enum class LanguagePreferenceKind { System, Culture };
struct LanguagePreference { LanguagePreferenceKind kind; std::wstring culture; };

HRESULT ApplyLanguagePreference(const LanguagePreference& preference) noexcept;
```

An empty value or `system` selects `LanguagePreferenceKind::System` (the Windows preferred UI-language chain); any other value selects that concrete culture.

### Lookup order

For a localized resource the manager tries the selected culture chain in satellite DLLs first, then falls back to the owner module's embedded English:

```text
selected culture satellite  →  (chain entries)  →  embedded English (owner module)
```

Key behaviors:

- Missing satellite DLLs, missing satellite IDs, and invalid/unavailable culture selections never block startup or UI creation. Embedded English is always the safety net.
- Non-string resources use the matching helper instead of a direct owner-module load: `LoadMenuResource`, `LoadAcceleratorsResource`, `FindLocalizedResourceHandle` / `FindLocalizedResource`, `LoadResourceBytes`, `LoadImageResource`.
- `FindLocalizedResourceHandle` returns both the `HINSTANCE` and the `HRSRC`. When you need `SizeofResource`/`LoadResource`/`LockResource`, use the `HINSTANCE` returned alongside the handle — never pair a satellite `HRSRC` with the embedded owner handle.
- Resource handles from `FindLocalizedResourceHandle` are transient. Copy or load the bytes immediately; do not cache the handle across a language change or owner unregister.
- Dialogs go through the resource-aware Win32 callback helpers so the satellite template is used when present, while the dialog is still created with the embedded owner `HINSTANCE` (so executable/plugin-registered child window classes keep resolving).
- Top-level menus are loaded explicitly via `LoadMenuResource`; window classes must not rely on `lpszMenuName` for localizable menus, because class-template loading bypasses the selected satellite.
- `DiscoverAvailableCultures()` enumerates the cultures available on disk (used to populate the language picker).

### Runtime language changes

A language change runs on the UI thread and must re-apply the preference, rebuild already-created top-level menu handles from `LoadMenuResource`, reset cached submenu handles, rebuild the dynamic theme/plugin menu sections, and resync the DxUi menu model. Modal dialogs already open may keep their current language until reopened.

## Language-neutral strings

Some resources are stable, non-translatable tokens (language autonyms in the language picker, product/protocol/brand names, file-format identifiers, keyboard glyphs, sample technical paths, placeholder-only layout skeletons, and technical formats such as hex/HRESULT skeletons). These stay resources, but they are owned only by the embedded English module and must never be duplicated into satellites.

- Load them explicitly with `LoadEmbeddedStringResource(ownerInstance, id)` or `FormatEmbeddedStringResource(ownerInstance, id, ...)`.
- Plugin code passes the plugin `HINSTANCE` (for example `g_hInstance`). Passing `nullptr` loads from the main executable and is correct only for main-app resources.
- This bypass of satellites is intentional and must be visible at the call site (that is why it uses a different helper than `LoadStringResource`).
- The full inventory is owner-scoped and documented inside `Tools/Tests/ResourceLocalizationContracts.Tests.ps1`. Ordinary UI words, sentences, command labels, tooltips, and grammar-bearing formats stay localizable even when the current translation happens to match English.

## Positional placeholder rules

Formatted resource strings use `std::format` syntax and must use positional placeholders so translators can reorder arguments for grammar.

| Rule | Allowed | Forbidden |
| --- | --- | --- |
| Index every placeholder | `{0}`, `{1:08X}` | bare `{}`, unindexed `{:08X}` |
| Use `std::format`, not printf | `0x{0:08X}: {1}` | `%s`, `%d`, `%08X` |
| Source introduces indexes in order | `Extracted {0} from {1}.` | `Extracted {1} from {0}.` (source) |

Detailed rules:

- Source (embedded English) strings introduce placeholders in argument order: `{0}`, then `{1}`, then `{2}`, with no skipped indexes. The `FormatStringResource(...)` argument list follows that same order.

  ```rc
  STRINGTABLE
  BEGIN
  IDS_FMT_HRESULT_DETAILS "0x{0:08X}: {1}"
  IDS_ARCHIVE_DONE        "Extracted {0} entries from {1}."
  END
  ```

  ```cpp
  auto msg = FormatStringResource(nullptr, IDS_ARCHIVE_DONE, entryCount, archivePath);
  ```

- Translated satellite strings may reorder placeholders for grammar, but must use the exact same placeholder tokens as the source: no added, dropped, duplicated, renumbered, or respecified placeholders. A translation of `{0} file{1:s}: {2} selected` may move `{2}`, but may not invent `{3}` or change `{2}` to `{2:s}`.
- Never treat a resource string as a printf format string. Use `FormatStringResource(...)` with positional placeholders and avoid C4774 suppressions.
- If a formatted string has invalid `std::format` syntax, the helper returns an empty fallback and logs the failing resource ID and the format error once. `std::bad_alloc` stays fatal and is not swallowed.

## String-id ranges

IDs are defined in the owner's `Resource.h`. Two conventions are worth knowing for the main app:

- **Command labels have two forms.** Full display names use `IDS_CMD_*` (menus, Preferences, shortcut lists). Short display names use `IDS_CMD_SHORT_BASE + IDS_CMD_*` for compact surfaces such as the function bar; every command must provide one.
- **`IDS_CMD_SHORT_BASE` is 20000**, and resource IDs `20000..21999` are reserved for command short labels so the arithmetic mapping cannot collide with unrelated strings. Do not reuse those IDs for anything else.

```c
// RedSalamander/Resource.h
#define IDS_CMD_SHORT_BASE 20000   // short labels = IDS_CMD_SHORT_BASE + IDS_CMD_* id
// IDs 20000-21999 reserved for command short labels.
```

The mapping is consumed in `RedSalamander/CommandRegistry.cpp` via `IDS_CMD_SHORT_BASE` and a `+ 1999` upper bound.

## Workflow: adding a string

1. **Pick an ID** in the owner's `Resource.h` (do not reuse an existing value or step into the `20000..21999` short-label range unless you are adding a command short label).
2. **Add the English text** to the owner's embedded `.rc` `STRINGTABLE`. Use positional placeholders for any variable text.
3. **Add a command short label too** if the new string is a command full name: define `IDS_CMD_SHORT_<NAME>` as `IDS_CMD_SHORT_BASE + <full id>`.
4. **Load it in code** with `LoadStringResource(id)` / `FormatStringResource(id, ...)` (or `MessageBoxResource`). For a stable non-translatable token, use `LoadEmbeddedStringResource`/`FormatEmbeddedStringResource` with the owning `HINSTANCE` and add the ID to the language-neutral inventory in `ResourceLocalizationContracts.Tests.ps1`.
5. **Mirror it into every existing satellite** for that owner under `Lang/<culture>/<Owner>-<culture>.rc` (unless the ID is language-neutral). Each non-neutral source ID must exist in each satellite, or the parity gate fails.
6. **Run the contract test** (see the gate section) after editing any braced/formatted string.

## Workflow: adding a culture

1. **Create the folder** `<Owner>/Lang/<culture>/` for each owner you are translating (start with the main app; plugins can follow incrementally).
2. **Add the satellite project** `<Owner>-<culture>.vcxproj`: a resource-only `DynamicLibrary` with the `<IsLanguageResourceProject>true</IsLanguageResourceProject>`, `<ResourceOwnerName>`, and `<ResourceCulture>` properties, compiling only the satellite `.rc`. Add it to `RedSalamander.sln`.
3. **Author the satellite `.rc`**: start with `#pragma code_page(65001)`, include the owner's `Resource.h`, declare the right `LANGUAGE LANG_*, SUBLANG_*`, and translate every non-neutral source ID. Keep placeholder tokens identical to the source.
4. **Expose the culture** so it appears in the language picker — `DiscoverAvailableCultures()` enumerates on-disk satellites; the persisted value is `ui.language` (for example `fr-FR`).
5. **Build and run the contract test** to confirm parity and placeholder safety.

### RedConfigure satellite generation

`RedConfigure.exe` can author satellite `.rc` files instead of hand-editing them. It discovers resource owners in the workspace, parses the embedded source `STRINGTABLE` and any existing target, merges them, and writes a satellite `.rc` for a chosen owner and culture (see `RedConfigure/Localization/RcWriter.h` / `RcWriter.cpp`, with `MergeStringTables` and `BuildSatelliteRcStringTable`). Generated files:

- keep resource text UTF-16/resource-compiler safe and use deterministic ordering,
- preserve positional `std::format` placeholders, and
- are blocked from export when a translation introduces bare `{}`, unindexed specs such as `{:08X}`, printf-style placeholders, or a placeholder set that does not match the source exactly.

Resource forms RedConfigure can parse but cannot safely rewrite yet stay visible as inventory and are not written silently.

The workbench supports ordered/pinned culture columns, rectangular TSV copy/paste, accelerator and placeholder inspection, preview-first batch changes, global validation, and explicit Review & Export. Generated satellite files are written through a sibling temporary file and reparsed after replacement. `RedConfigureTests` also compiles a generated fixture with installed Windows SDK `rc.exe`; missing SDK tooling is the only permitted skip condition for that compiler check.

## The contract gate (ResourceLocalizationContracts.Tests.ps1)

`Tools/Tests/ResourceLocalizationContracts.Tests.ps1` is a Pester gate that scans every `.rc` in the repo (excluding `.build`, `packages`, `.claude`). Run it whenever you add or edit braced/formatted resource strings; it also runs as part of `Tools/Run-AllTests.ps1 -Suite Full`.

It enforces four contracts:

| Contract | What it checks |
| --- | --- |
| Placeholder safety | No bare `{}`, no unindexed `{:...}`, no printf `%s`/`%d`; source strings index placeholders from `0` with no skips and in first-use order. |
| Translation placeholder parity | Each satellite string's indexed-placeholder signature matches its source string exactly (reordering allowed, renumber/add/drop/respec not). |
| Language-neutral inventory | Each documented neutral ID exists in the embedded owner `.rc`, is absent from every satellite for that owner, and is loaded only through the embedded helpers (plugins must pass the plugin `HINSTANCE`, not `nullptr`). |
| Satellite ID parity | Every source ID exists in each satellite (except documented neutral IDs), and no satellite carries an ID the source does not define. |

Run it directly with Pester, for example:

```powershell
Invoke-Pester -Path .\Tools\Tests\ResourceLocalizationContracts.Tests.ps1
```

## See also

- [Developer Guide](../DeveloperGuide.md) — Localization, Resources, Build & Test Infrastructure overview
- [Themes.md](../Themes.md) — the dynamic theme menu sections rebuilt on language change
- [Preferences.md](../Preferences.md) — the General → Display language setting
- `Specs/Core/Core_Localization.md` — normative specification
- `.github/skills/localization/SKILL.md` — patterns and helper reference


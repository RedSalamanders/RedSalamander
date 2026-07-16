---
name: theming
description: Theme system and color key management for RedSalamander. Use when adding theme colors, modifying AppTheme, working with theme.json5 files, or updating theme-related specifications.
metadata:
  author: RedSalamander
  version: "2.0"
---

# Theming System

## When Modifying Theme Color Keys

Update **ALL** of these locations:

1. `Specs/SettingsStore.schema.json` - Validation + IntelliSense
2. `Specs/Core/Core_SettingsStore.md` - Durable format and resolution contract
3. `Common/ThemeExpression.h` / `Common/Common/ThemeExpression.cpp` - Shared source parser, resolver, and compiled dynamic evaluator
4. `Common/ThemeDefinitionIo.h` / `Common/Common/ThemeDefinitionIo.cpp` - Shared JSON5 reader/writer
5. `RedSalamander/AppTheme.h` / `AppTheme.cpp` - Defaults
6. Runtime override mappings in RedSalamander, RedSalamanderMonitor, and hot reload
7. `Specs/Themes/*.theme.json5` - Shipped themes
8. Component specs and tests referencing the token
9. Public documentation in `docs/Themes.md`, `docs/Preferences.md`, `docs/UserGuide.md`, and `docs/DeveloperGuide.md`
10. Regenerated `docs/res/theme-controls-*.png` galleries and the theme distribution/gallery contract

## Theme File Format (JSON5)

Location: `Specs/Themes/` or `Themes/` next to exe

```json5
{
  "formatVersion": 2,
  "id": "user/example",
  "name": "Example",
  "baseThemeId": "builtin/dark",
  "palette": {
    "background": "#1E1E1E",
    "foreground": "#D4D4D4",
    "accent": "#007ACC",
    "selection": "blend(palette.background,palette.accent,25%)"
  },
  "colors": {
    "app.accent": "ref(palette.accent)",
    "window.background": "ref(palette.background)",
    "folderView.itemBackgroundSelected": "ref(palette.selection)"
  }
}
```

`formatVersion: 2` is mandatory. Do not add a legacy reader, accept a missing version, or flatten expressions on export. Palette names cannot contain `.`. Expressions are non-nested and use named palette or semantic references.

Use only the functions documented in `Specs/Core/Core_SettingsStore.md`. Parse and resolve with the shared Common APIs; do not add a consumer-local parser. Paint-time sources must be compiled at theme resolution and evaluated with `EvaluateDynamicThemeColor`; paint code must not parse, allocate, lock, or perform I/O.

## Runtime Integration

```cpp
Common::Settings::ResolvedThemeColors resolved;
auto context = Common::Settings::MakeSystemThemeResolutionContext(effectiveDark);
RETURN_IF_FAILED(Common::Settings::ResolveThemeDefinition(theme, context, resolved));
ApplyThemeOverrides(appTheme, resolved.colors);
```

Provide a base-color lookup when expressions may reference inherited semantic tokens. Keep `builtin/rainbow` as the application-wide Rainbow base; a static semantic override suppresses Rainbow only for that token. High Contrast always wins.

## Best Practices

- Use semantic token names in `colors` and descriptive reusable names in `palette`
- Always provide fallback colors
- Test in both light and dark modes
- Validate contrast ratios for accessibility
- Add deterministic tests for missing references, cycles, event refresh, and seeded dynamic output
- Add exact upstream license text and update `Specs/Themes/THIRD-PARTY-NOTICES.md` for third-party palettes
- Regenerate the complete DxUi control gallery after adding, removing, renaming, or materially changing a shipped theme; keep separate deterministic Rainbow Light and Rainbow Dark galleries and remove stale gallery files
- Preserve authored intent in editing operations: palette rename rewrites references atomically, delete reports/block dependents, repeated literals can become one palette reference, and batch transforms create palette/function recipes instead of flattened hex values
- For resolver, preview, or paint-time changes, run `theme_v2_runtime_resolution_and_dynamic_perf`, archive the run under `Specs/TestRuns/`, and report p50/p95/max for `theme.resolve_us` and `theme.dynamic.evaluate_us`

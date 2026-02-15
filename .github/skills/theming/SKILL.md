---
name: theming
description: Theme system and color key management for Red Salamander. Use when adding theme colors, modifying AppTheme, working with theme.json5 files, or updating theme-related specifications.
metadata:
  author: DualTail
  version: "1.0"
---

# Theming System

## When Modifying Theme Color Keys

Update **ALL** of these locations:

1. `Specs/SettingsStore.schema.json` - Validation + IntelliSense
2. `Specs/SettingsStoreSpec.md` - Key documentation
3. `RedSalamander/AppTheme.h` / `AppTheme.cpp` - Defaults
4. `RedSalamander/RedSalamander.cpp` (`ApplyThemeOverrides`) - Mapping
5. `Specs/Themes/*.theme.json5` - Shipped themes
6. Component specs referencing the token

## Theme File Format (JSON5)

Location: `Specs/Themes/` or `Themes/` next to exe

```json5
{
    "name": "Dark Theme",
    "colors": {
        "background": "#1E1E1E",
        "foreground": "#D4D4D4",
        "accent": "#007ACC",
        "selection": "#264F78",
        "border": "#3C3C3C"
    }
}
```

## AppTheme Integration

```cpp
// AppTheme.h
struct ThemeColors {
    D2D1_COLOR_F background;
    D2D1_COLOR_F foreground;
    D2D1_COLOR_F accent;
};

// Load and apply
void ApplyThemeOverrides(ThemeColors& colors, const json& theme) 
{
    if (theme.contains("colors")) {
        auto& c = theme["colors"];
        if (c.contains("background")) {
            colors.background = ParseColor(c["background"]);
        }
        // Always provide fallbacks for missing keys
    }
}
```

## Best Practices

- Use semantic color names (`selection` not `blue1`)
- Always provide fallback colors
- Test in both light and dark modes
- Validate contrast ratios for accessibility

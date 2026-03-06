# Themes

RedSalamander is fully theme-aware (menus, panes, file operations UI, and viewers).

![Themes in Preferences](res/preferences-themes.png)

## Quick theme switching

Use **View → Theme**:

- **System**
- **Light**
- **Dark**
- **Rainbow**
- **High Contrast (App)**

Note:

- **High Contrast (System)** is shown as an indicator when Windows high-contrast is enabled.

## Theme sources

RedSalamander supports theme files:

- `Themes\*.theme.json5` next to `RedSalamander.exe`

It also supports user themes stored in your settings file as `user/*` themes.

## Preferences → Themes

- Select a built-in, file-based, or user theme.
- Create a new user theme or duplicate an existing theme into a user theme.
- Edit explicit color overrides while inherited values remain visible.
- Preview a theme immediately with **Apply Temporarily**.
- Load/save `*.theme.json5` theme files.

## Notes

- Cancelling Preferences after a temporary preview restores the previously applied theme.
- The main **View → Theme** menu is the fastest way to switch between the standard built-in themes.

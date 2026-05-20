# Themes

RedSalamander is theme-aware across the main chrome, panes, navigation bars, file operations UI, dialogs, and viewers.

![Themes in Preferences](res/preferences-themes.png)

## Control galleries

The generated control galleries show the shared DxUi controls rendered under each built-in theme and each shipped `Specs/Themes/*.theme.json5` theme.

### Button State Contrast

The button contrast audit shows standard and primary buttons across idle, hover, pressed, keyboard-focus, and disabled states. Enabled text is marked against WCAG normal-text thresholds; disabled controls are shown but marked exempt.

![Button state contrast audit](res/theme-button-states-after-fix.png)

### Light

![Light controls](res/theme-controls-light.png)

### Dark

![Dark controls](res/theme-controls-dark.png)

### Rainbow

![Rainbow controls](res/theme-controls-rainbow.png)

### High Contrast

![High Contrast controls](res/theme-controls-high-contrast.png)

### Forest Mist

![Forest Mist controls](res/theme-controls-forest-mist.png)

### Neon Tokyo

![Neon Tokyo controls](res/theme-controls-neon-tokyo.png)

### Paper & Ink

![Paper & Ink controls](res/theme-controls-paper-ink.png)

### Retro Terminal

![Retro Terminal controls](res/theme-controls-retro-terminal.png)

### Solar Flare

![Solar Flare controls](res/theme-controls-solar-flare.png)

### Ugly

![Ugly controls](res/theme-controls-ugly.png)

## Quick theme switching

Use **View → Theme**:

- **System**
- **Light**
- **Dark**
- **Rainbow**
- **High Contrast (App)**

Note:

- **High Contrast (System)** is shown as an indicator when Windows high-contrast is enabled.

## Theme examples

The following screenshots were captured from the running application using different built-in themes.

![Light theme](res/theme-light.png)

![Dark theme](res/theme-dark.png)

![Rainbow theme](res/theme-rainbow.png)

![High Contrast App theme](res/theme-high-contrast-app.png)

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

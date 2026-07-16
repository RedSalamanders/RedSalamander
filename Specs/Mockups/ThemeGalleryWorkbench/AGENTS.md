# Prototype Instructions

Run the local server yourself and open the preview in the browser available to this environment. Do not give the user server-start instructions when you can run it.

Before making substantial visual changes, use the Product Design plugin's `get-context` skill when the visual source is unclear or no longer matches the current goal. When the user gives durable prototype-specific design feedback, preferences, or decisions, record them in `AGENTS.md`.

When implementing from a selected generated mock, treat that image as the source of truth for layout, component anatomy, density, spacing, color, typography, visible content, and hierarchy.

## Product decisions

- Theme comparison uses one large, faithful RedSalamander application shell rather than a grid of distant miniature windows.
- Do not duplicate plugin preview surfaces when plugins share the same theme anatomy; the mockup focuses on the host application's real dual-pane UI.
- The application preview follows the repository screenshots and UI contracts: title/menu rows, two NavigationViews, two FolderViews, a 6 DIP-style splitter, pane status, and the function-key bar.
- Built-in Rainbow is inspectable as explicit Light and Dark variants.
- Multiple selection is the default sample. In Rainbow, every selected row receives its own stable deterministic tint, while the current item keeps the stronger focus border.
- Dynamic theming functions must be demonstrated through one interactive function lab attached to the real application preview. `seededChoice`, `systemColor`, `perceptualTone`, `ensureContrast`, and `harmonize` each need a visible target, bounded controls, authored expression, evaluation phase, before/after swatches, and a live application effect.
- Catppuccin uses the same semantic mapping in every flavor: Lavender for active/focus accent, Blue for actions and links, and Overlay2 at 26% over Base for selection. Mauve remains available in the palette but is not the default global accent.
- Dracula and Catppuccin entries visibly identify their palette license family, and implementation must ship the required license/attribution text.

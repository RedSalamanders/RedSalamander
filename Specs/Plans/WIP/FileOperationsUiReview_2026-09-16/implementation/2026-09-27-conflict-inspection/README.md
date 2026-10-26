# Conflict inspection: owned foreground Debug witness

The single governed Commands case passes 1/0/0 on x64 Debug, run `20260927T151422Z-87984-01e82b162f8d41ada5bd932ec31fbe5b`, with zero disk-audit findings. These 28 images are byte-exact application captures from its deterministic harness. [Provenance and image hashes](provenance.json) bind the full Debug build; [raw results and retained failed attempt](../../../../../TestRuns/4cb089111a23/Validation/2026-09-27_migration-transition/README.md) record qualification limits.

All twelve combinations of pane (left/right), role (source/destination) and invocation (UIA/Space/Enter) verify exact selected item and pane focus, hide the popup without deciding the conflict, then return to the same unresolved card, scope, scroll and invoking control. The focused case also exercises missing-item feedback, short-body action reachability and stale/focus-change refusal. The existing visible warning, interactive lease and focus/cursor restoration scope guard are retained. This is a focused Debug witness; Release, native DPI transitions, external assistive technology and the Full runner remain separate gates.

A cheaper-model reviewer inspected all 21 SHA-unique images; root checked six representative originals and corrected the review of the generic `selected__fr-FR.png`: it depicts popup feedback, not the file pane. The twelve specifically named selected-pane images are the pane-selection evidence. Paths intentionally scroll inside the comparison body; boundary clipping in a scrolled view is not evidence that its complete value fits without scrolling. The short-body capture establishes reachability of the focused action, not complete-path visibility. Buttons and the prominent `payload.txt` subject remain readable in the reviewed originals.

## Selection and return pairs

| Pane / role / invocation | Selected file pane | Returned unresolved card |
|---|---|---|
| left / source / uia | [Selected](selected__left__source__uia__fr-FR.png) | [Returned](returned__left__source__uia__fr-FR.png) |
| left / source / space | [Selected](selected__left__source__space__fr-FR.png) | [Returned](returned__left__source__space__fr-FR.png) |
| left / source / enter | [Selected](selected__left__source__enter__fr-FR.png) | [Returned](returned__left__source__enter__fr-FR.png) |
| left / destination / uia | [Selected](selected__left__destination__uia__fr-FR.png) | [Returned](returned__left__destination__uia__fr-FR.png) |
| left / destination / space | [Selected](selected__left__destination__space__fr-FR.png) | [Returned](returned__left__destination__space__fr-FR.png) |
| left / destination / enter | [Selected](selected__left__destination__enter__fr-FR.png) | [Returned](returned__left__destination__enter__fr-FR.png) |
| right / source / uia | [Selected](selected__right__source__uia__fr-FR.png) | [Returned](returned__right__source__uia__fr-FR.png) |
| right / source / space | [Selected](selected__right__source__space__fr-FR.png) | [Returned](returned__right__source__space__fr-FR.png) |
| right / source / enter | [Selected](selected__right__source__enter__fr-FR.png) | [Returned](returned__right__source__enter__fr-FR.png) |
| right / destination / uia | [Selected](selected__right__destination__uia__fr-FR.png) | [Returned](returned__right__destination__uia__fr-FR.png) |
| right / destination / space | [Selected](selected__right__destination__space__fr-FR.png) | [Returned](returned__right__destination__space__fr-FR.png) |
| right / destination / enter | [Selected](selected__right__destination__enter__fr-FR.png) | [Returned](returned__right__destination__enter__fr-FR.png) |

## Additional states

- [Popup feedback before focus restoration](selected__fr-FR.png)
- [Popup feedback with restored source-action focus](returned__fr-FR.png)
- [Missing item: no selection claimed](missing-item__fr-FR.png)
- [Constrained body: focused More action remains visible](short-body-focused-action__fr-FR.png)

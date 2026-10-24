# Planned Command Backlog Implementation Plan

## Live Implementation Checklist

Status legend: `[todo]` not started, `[doing]` actively being changed or audited, `[review]` implemented and awaiting verification, `[done]` verified with archived evidence, `[blocked]` needs a decision or missing dependency.

- [done] Completed slice: Phase 6 `cmd/pane/shares`. RED build `.build/logs/msbuild-20260503_014506_541.log`; implementation build `.build/logs/msbuild-20260503_015901_797.log`; post-doc/resource rebuild `.build/logs/msbuild-20260503_020351_523.log`; focused GREEN behavior/perf run `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_020538` with `sharedDirectories.open_us` = 36,265 us and `rowCount` = 2; stale-label guard `implemented_menu_labels_not_todo` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_020528`.
  - [done] Shared Directories RED selftest and perf artifact expectation are patched.
  - [done] Shared Directories production dialog/model/debug provider/command routing implementation builds cleanly in `.build/logs/msbuild-20260503_015901_797.log`.
  - [done] Shared Directories focused GREEN behavior/perf selftest passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_020538` with `sharedDirectories.open_us` = 36,265 us and `rowCount` = 2.
  - [done] Shared Directories docs/spec/localization cleanup and stale-label guard passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_020528`.
  - [done] Next slice: Pack and Unpack implementation.
    - [done] Archive slice checkpoint A: refresh the live plan checklist and split `cmd/pane/pack` / `cmd/pane/unpack` into test, implementation, verification, perf, and spec closeout steps.
    - [done] Archive slice checkpoint B: RED command selftest `cmd_pane_archive_pack_unpack_zip_roundtrip_and_validation` is patched for ZIP pack/unpack round trip, overwrite behavior, invalid destination/path validation, unsupported provider feedback, and archived perf artifact expectations. RED build `.build/logs/msbuild-20260503_021445_724.log` fails on the missing archive command debug/implementation surface.
    - [done] Archive slice checkpoint C: implement the ZIP archive engine, command routing, debug automation hooks, localized feedback, and menu/spec cleanup. Implementation build passed in `.build/logs/msbuild-20260503_023247_097.log`; post-cleanup rebuild passed in `.build/logs/msbuild-20260503_024049_432.log`.
      - [done] Deterministic stored-ZIP pack/unpack helpers are patched with CRC, sorted entries, overwrite validation, traversal-safe extraction, and perf counters; implementation build passed in `.build/logs/msbuild-20260503_023247_097.log`.
      - [done] Localized archive feedback strings are patched in English and French resources.
      - [done] Command methods and test debug automation result plumbing are patched; implementation build passed in `.build/logs/msbuild-20260503_023247_097.log`.
    - [done] Archive slice checkpoint D: run focused build, command selftests, stale-label guard, and perf evidence archival serially.
      - [done] Implementation build passed in `.build/logs/msbuild-20260503_023247_097.log`.
      - [done] Focused archive behavior/perf selftest passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_023445` with `archive.pack_us` = 81,016 us, `archive.unpack_us` = 74,722 us, `entryCount` = 16, and `bytesProcessed` = 23.
      - [done] Stale archive menu/spec `[todo]` labels are removed and Pack/Unpack are added to the stale-label guard. Post-cleanup build `.build/logs/msbuild-20260503_023859_420.log` caught the guard array-size update; post-fix rebuild passed in `.build/logs/msbuild-20260503_024049_432.log`.
      - [done] Stale-label guard `implemented_menu_labels_not_todo` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_024221`.
      - [done] Final focused archive behavior/perf rerun passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_024236` with `archive.pack_us` = 94,444 us, `archive.unpack_us` = 76,226 us, `entryCount` = 16, and `bytesProcessed` = 23.
      - [done] Fresh closeout archive behavior/perf rerun passed after final rebuild in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_024835` with `archive.pack_us` = 91,898 us, `archive.unpack_us` = 154,142 us, `entryCount` = 16, and `bytesProcessed` = 23.
    - [done] Archive slice checkpoint E: durable archive behavior is merged into `Specs/UI/UI_CommandMenuKeyboard.md`, `Docs/UserGuide.md`, `Docs/FileOperations.md`, `Docs/Plugins.md`, `Docs/Todo.md`, and `Specs/Testing/Testing_TestCoverage.md`; `cmd/pane/pack` and `cmd/pane/unpack` are marked complete.
- [done] Previous slice `cmd/pane/listOpenedFiles` is verified: RED build `.build/logs/msbuild-20260503_010701_401.log`; GREEN build `.build/logs/msbuild-20260503_012931_250.log`; post-cleanup build `.build/logs/msbuild-20260503_013820_113.log`; focused behavior/perf rerun `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_014045` passed with `rowCount` = 3 and `listOpenedFiles.open_us` = 93,454 us; stale-label guard `implemented_menu_labels_not_todo` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_014104`.
- [done] Previous slice `cmd/pane/makeFileList` is verified: RED build `.build/logs/msbuild-20260503_002510_245.log`; GREEN build `.build/logs/msbuild-20260503_005611_159.log`; fresh settings/schema run `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_005737`; fresh behavior/perf run `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_005745` with JSON/CSV/text output coverage, `entryCount` = 4 for JSON, test-measured JSON command duration = 95,435 us, and command metrics `makeFileList.collect_us`, `makeFileList.generate_us`, `makeFileList.output_us`, `makeFileList.feedback_us`, and `makeFileList.total_us`; fresh stale-label guard `implemented_menu_labels_not_todo` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_005751`.
- [done] Previous slice `cmd/app/rereadAssociations` is verified: RED build `.build/logs/msbuild-20260503_000552_653.log`; RED focused run `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_000742`; GREEN build `.build/logs/msbuild-20260503_001755_942.log`; fresh GREEN focused behavior/perf run `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_001923` with `rereadAssociations.total_us` = 323,179 us, association icon cache seed count = 50, and viewer/editor/User Menu action counts = 1/1/1; fresh stale-label guard `implemented_menu_labels_not_todo` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_001929`.
- [done] Foundation guard: `cmd/app/viewWidth` remains the existing splitter view-width command.
- [done] Foundation guard: `cmd/pane/editWidth` is retired from command registry, resources, menus, shortcut defaults, and specs; only regression assertions may reference it.
- [done] Foundation command: `cmd/pane/alternateEdit` is registered, localized, in the Files menu, and bound to `Ctrl+Shift+F4`.
- [done] Foundation command: `cmd/app/theme/selectNext` and `cmd/app/theme/selectPrev` are registered, localized, in the Theme menu, bound to function-bar shortcuts, and covered by command selftests.
- [done] Phase 0.1: Audit all planned/not-implemented command surfaces from spec, menus, docs, registry, shortcut defaults, and fallback dispatch.
- [done] Phase 0.2: Expand this checklist so every missing command family, dynamic command pattern, settings page, schema update, test, perf scenario, and spec closeout item is visible. Rechecked against command registry, menu `[todo]` labels, fallback dispatch, shortcut defaults, and spec command ids during Phase 2 continuation.
- [done] Phase 0.3: Add command inventory selftests that fail when `[todo]` menus/spec markers, registry entries, shortcut defaults, and command handlers drift. RED: `implemented_menu_labels_not_todo` failed on Right Hot Paths `[todo]` in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_125931`; GREEN: passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_130325`.
- [done] Phase 0.4: Remove stale `[td]` / `[todo]` labels only for commands proven implemented by focused selftests.
- [done] Phase 1: Add a single command execution layer for shortcut commands, pane-scoped commands, and parameterized command ids. RED: `theme_cycle_commands` failed on `cmd/app/theme/select/<themeId>` in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_130626`; RED: `registry_integrity` failed on missing parameter canonicalization in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_131418`; GREEN: `registry_integrity` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_133014`, `theme_cycle_commands` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_133019`, `generic_status_bar_command_routes_active_pane` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_133030`.
- [done] Phase 2: Implement settings-driven View, Alternate View, Edit, Alternate Edit, View With, Edit With, Edit New, and the Viewers/Editors preferences contract. Final slice GREEN: `cmd_preferences_dialog_viewers_editors_file_action_settings_apply` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_172731` with `preferences.ui.editors_layout_us` and `preferences.ui.viewers_layout_us`; docs/spec closeout updated `Specs/Core/Core_SettingsStore.md`, `Specs/UI/UI_CommandMenuKeyboard.md`, and user docs.
- [done] Phase 3: Implement shell and item utility commands: current directory context menu, open security, change attributes, go to shortcut target, and shell templates.
- [done] Phase 4: Implement clipboard, command-line, quick search, and user menu commands.
- [done] Phase 5: Implement pane view option commands: file extensions, thumbnails, preview pane, filter bar, navigation bar, and generic status bar routing.
- [done] Phase 6: Implement Make File List, List Opened Files, Shared Directories, Reread Associations, Pack, and Unpack.
- [done] Phase 7: Update authoritative specs, docs, localization, tests, perf evidence, and close the plan into `Specs/Plans/Done/`.

### Detailed Live Work Items

Phase 0 inventory and stale-marker cleanup:

- [done] Runtime menu guard: implemented commands must not carry `[todo]` in the loaded main menu.
- [done] Spec cleanup: remove stale `[td]` / `*(planned)*` markers for implemented `cmd/pane/permanentDelete`, `cmd/pane/sort/attributes`, `cmd/pane/selection/selectSameName`, `cmd/pane/selection/unselectSameName`, `cmd/app/plugins/manage`, `cmd/pane/selectFileSystemPlugin`, `cmd/pane/windowMenu`, `cmd/app/toggleFunctionBar`, `cmd/app/toggleMenuBar`, and left/right status bar menu entries.
- [done] Menu cleanup: remove stale `[todo]` from Right > Go to > Hot Paths in English and French resources.
- [done] Docs cleanup: remove stale `Edit width` wording from `Docs/Todo.md`; keep `cmd/app/viewWidth` as the implemented splitter-width command.
- [done] Docs inventory: add Mouse preference page work as a separate non-command backlog item, because it is planned but not part of the command backlog.
- [done] Command inventory guard: every command in `CommandRegistry` is covered by `registry_integrity`, planned-command canonicalization, or dynamic/parameterized command-family coverage.
- [done] Fallback inventory guard: registered WM commands that still reach the localized not-implemented path are constrained by the planned-command backlog and covered by `generic_status_bar_command_routes_active_pane` plus the stale-label guard.
- [done] Shortcut inventory guard: default shortcuts that target planned commands remain deliberate and are covered by `shortcut_defaults_mapping`.

Phase 1 command execution layer:

- [done] Add a single command-id execution entrypoint used by shortcut dispatch and reusable from WM-command handlers.
- [done] Preserve existing dynamic menu maps for path, theme, and plugin menu ids until command parameters carry equivalent payloads.
- [done] Canonicalize parameterized command families: theme select, plugin configure/toggle, file-system plugin select, viewer/editor choices, user-menu items, shell templates, navigation paths, known folders, go-drive-root, hot-path, and set-hot-path.
- [done] Add selftests for command-id execution, fallback behavior, active-pane targeting, and no silent no-op. RED: `generic_status_bar_command_routes_active_pane` failed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_132052`; GREEN: passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_133030`.

Phase 2 settings-driven View/Edit family:

- [done] Settings model and schema for viewer actions: primary view, alternate view, view-with list, internal viewer actions, external program actions, enabled state, ordering, extension filters, computer filters, default action ids, computer-specific overrides, executable path, arguments, and working directory. RED: build failed on missing `FileActionDefinition`, `FileActionKind`, and `settings.viewers` in `.build/logs/msbuild-20260502_133354_821.log`; GREEN: `settings_store_view_edit_actions_roundtrip` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_134409`.
- [done] Settings model and schema for editor actions: primary edit, alternate edit, edit-with list, edit-new defaults, enabled state, ordering, extension filters, computer filters, default action ids, computer-specific overrides, executable path, arguments, and working directory. RED: build failed on missing `settings.editors` in `.build/logs/msbuild-20260502_133354_821.log`; GREEN: `settings_store_view_edit_actions_roundtrip` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_134409`.
- [done] Shared macro engine for external actions: path, full path, path plus filename, filename, selected-paths file, opposite pane path, computer name, and literal escaping. RED: build failed on missing `FileActionLauncher.h` in `.build/logs/msbuild-20260502_134637_593.log`; GREEN: `file_action_external_launch_plan_macros` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_135051`.
- [done] Dry-run launch-plan builder selftests before any real process launch path changes. RED: `file_action_external_launch_plan_macros` could not build until `FileActionLauncher` existed; GREEN: dry-run launch-plan macro validation passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_135051`.
- [done] Shared external-program process launcher primitive with hidden/waitable selftest coverage. RED: build failed on missing `LaunchOptions`, `LaunchResult`, and `LaunchExternalPlan` in `.build/logs/msbuild-20260502_141913_735.log`; GREEN: `file_action_external_launch_starts_process` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_142458` with process launch/wait correctness coverage. Command-family coverage is tracked separately below for View, Edit, and User Menu.
- [done] Shared file-action resolver for primary/alternate extension mappings, default fallbacks, computer overrides, disabled actions, and picker ordering. RED: build failed on missing `ResolveActionForFile`, `ActionRole`, and `CollectApplicableActions` in `.build/logs/msbuild-20260502_135217_192.log`; GREEN: `file_action_resolution_uses_extension_computer_and_fallback` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_135613`.
- [done] Command wiring for `cmd/pane/alternateView`: add alternate-view request role, route main/compare WM commands, and resolve configured viewer-plugin actions from settings. RED: `alternate_view_uses_configured_viewer_action` failed because no viewer instance opened in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_140520`; GREEN: passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_140942`.
- [done] Command wiring for primary `cmd/pane/view`: resolve configured primary viewer actions before legacy extension mappings, including internal and external action coverage. Coverage-only GREEN: `view_command_uses_configured_primary_viewer_action` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_151024` with `fileaction.external.launch_us` = 30,114 us.
- [done] Direct `cmd/pane/viewWith/<viewerId>` shortcut dispatch to a configured viewer action. RED: `view_with_parameterized_command_uses_configured_viewer_action` failed because no viewer instance opened in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_141351`; GREEN: passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_141714`.
- [done] `cmd/pane/viewWith` picker UI and dynamic menu population: opens the View With popup, shows applicable configured viewer actions for the focused file, dispatches the generated menu id, launches the selected action, and records process-launch perf evidence. RED: `view_with_menu_populates_applicable_viewer_actions` failed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_144745`; GREEN: passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_145609` with `fileaction.viewwith.menu_populate_us` = 93 us and `fileaction.external.launch_us` = 32,718 us.
- [done] `cmd/pane/viewWith/<viewerId>` external-program actions plus disabled/unavailable action reporting. External-program launch for valid direct action ids is covered. RED: `view_with_parameterized_command_launches_external_viewer_action` failed because the external action was not launched in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_142855`; GREEN: passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_143246` with `fileaction.external.launch_us` = 33,042 us. RED: `view_with_disabled_action_reports_alert` failed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_155744` because a disabled configured viewer action returned silently without a pane alert; GREEN: passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_160039`.
- [done] `cmd/pane/alternateView` external-program actions, selected-paths-file macro lifecycle, and localized "no alternate viewer configured" feedback. External-program action coverage is green: `alternate_view_launches_external_viewer_action` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_152908` with `fileaction.selected_paths_file.write_us` = 431 us, `fileaction.external.launch_us` = 31,422 us, and async `fileaction.selected_paths_file.cleanup_us` = 152 us. RED: `alternate_view_without_configured_action_reports_alert` failed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_160352` because Alternate View opened the regular viewer fallback instead of reporting that no alternate viewer is configured. GREEN: passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_160842` with `fileaction.feedback_us` = 619 us for `alternate-viewer-unavailable`.
- [done] View/Edit external launch-failure feedback: broken external viewer/editor actions show localized pane alerts with the action id, focused file, and HRESULT. RED: `view_with_launch_failure_reports_alert` failed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_162202` because the alert did not include the launch HRESULT. RED: `edit_with_launch_failure_reports_alert` failed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_162209` because the alert did not include the launch HRESULT. GREEN: `view_with_launch_failure_reports_alert` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_162543` with `fileaction.feedback_us` = 756 us for `viewer-launch-failed`. GREEN: `edit_with_launch_failure_reports_alert` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_162552` with `fileaction.feedback_us` = 549 us for `editor-launch-failed`.
- [done] View/Edit alert regression sweep: fresh focused runs passed for `view_with_disabled_action_reports_alert`, `alternate_view_without_configured_action_reports_alert`, `edit_with_disabled_action_reports_alert`, `alternate_edit_without_configured_action_reports_alert`, `view_with_launch_failure_reports_alert`, and `edit_with_launch_failure_reports_alert` in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_162708` through `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_162750`.
- [done] Selected-paths-file lifecycle for external viewer/editor/user-menu actions: create UTF-16 list files only when `{SelectedPathsFile}` is used, pass the path to macros, and delete/report cleanup deterministically. RED: `file_action_selected_paths_file_lifecycle` failed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_151532` because `{SelectedPathsFile}` still required a pre-supplied file. GREEN: passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_152334` with `fileaction.selected_paths_file.write_us` = 325 us, `fileaction.selected_paths_file.cleanup_us` = 180 us, and `fileaction.external.launch_us` = 77,300 us.
- [done] File-action regression sweep: `file_action_` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_152608` covering macro expansion, process launch, selected-paths lifecycle, and action resolution after the selected-paths contract update.
- [done] `cmd/pane/edit`, `cmd/pane/editWith`, `cmd/pane/editWith/<editorId>`, and `cmd/pane/alternateEdit` launch and unavailable-feedback basics. Direct `cmd/pane/editWith/<editorId>` external-program launch for a focused file is covered. RED: `edit_with_parameterized_command_launches_external_editor_action` failed because the external editor action was not launched in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_143629`; GREEN: passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_144005` with `fileaction.external.launch_us` = 31,266 us. Edit With picker dynamic menu population and generated-id routing is covered: RED `edit_with_menu_populates_applicable_editor_actions` failed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_150021`; GREEN passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_150301` with `fileaction.editwith.menu_populate_us` = 81 us and `fileaction.external.launch_us` = 32,441 us. Primary Edit focused-file behavior is fixed: RED `edit_command_uses_focused_primary_editor_action` failed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_153529` because Edit launched the selected file instead of the focused file; GREEN passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_153823` with `fileaction.selected_paths_file.write_us` = 406 us, `fileaction.external.launch_us` = 31,171 us, and async cleanup = 191 us. Alternate Edit external action coverage passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_153845` with `fileaction.external.launch_us` = 30,954 us. Edit sweep `edit_` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_153937`, covering direct Edit With, primary Edit, and dynamic Edit With menu. English/French menu labels and command spec markers are clean for implemented View/Edit commands: `implemented_menu_labels_not_todo` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_154639`, guarding primary View, Alternate View, Edit, Alternate Edit, View With, and Edit With. RED: `edit_with_disabled_action_reports_alert` failed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_161305` because disabled Edit With returned silently. GREEN: passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_161624` with `fileaction.feedback_us` = 537 us for `editor-unavailable`. RED: `alternate_edit_without_configured_action_reports_alert` failed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_161321` because Alternate Edit returned silently. GREEN: passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_161631` with `fileaction.feedback_us` = 1,006 us for `alternate-editor-unavailable`. Launch-failure feedback closed with `view_with_launch_failure_reports_alert` and `edit_with_launch_failure_reports_alert` passing in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_162543` and `2026-05-02_162552`.
- [done] `cmd/pane/editNew` dialog with file-name validation and Editor combo filtered by extension, computer name, settings, and executable availability. RED: build failed in `.build/logs/msbuild-20260502_163417_695.log` because `GetFolderViewEditNewPromptHandle`, `FolderViewEditNewPromptDebugSnapshot`, and related Edit New prompt debug APIs do not exist yet. GREEN build: `.build/logs/msbuild-20260502_164549_625.log`. GREEN focused run: `cmd_pane_editNew_prompt_filters_editor_combo_and_creates_file` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_164741` with `fileaction.editnew.editor_combo_us` = 156 us for one applicable editor, `fileaction.editnew.create_file_us` = 1,021 us, and `fileaction.external.launch_us` = 49,131 us.
- [done] `cmd/pane/editNew` validation, no-editor behavior, and unlaunchable-editor filtering. RED: `cmd_pane_editNew_prompt_filters_editor_combo_and_creates_file` failed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_165329` because the Editor combo included a matching external editor with no executable path (`fileaction.editnew.editor_combo_us` reported two actions). GREEN build: `.build/logs/msbuild-20260502_165435_788.log`. GREEN runs: `cmd_pane_editNew_prompt_filters_editor_combo_and_creates_file` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_165611` with one applicable editor and `fileaction.external.launch_us` = 38,168 us; `cmd_pane_editNew_prompt_rejects_invalid_and_existing_names` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_165644`; `cmd_pane_editNew_prompt_creates_file_without_applicable_editor` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_165650` with `fileaction.editnew.create_file_us` = 772 us and no external launch.
- [done] `cmd/pane/editNew` stale marker cleanup: English/French menus and `Specs/UI/UI_CommandMenuKeyboard.md` no longer mark Edit New as TODO. GREEN build: `.build/logs/msbuild-20260502_165821_284.log`; GREEN guard: `implemented_menu_labels_not_todo` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_170000`.
- [done] Viewers and Editors preferences UI plus settings schema/save/load roundtrip tests. RED: `cmd_preferences_dialog_viewers_editors_file_action_settings_apply` fails to build in `.build/logs/msbuild-20260502_170935_264.log` because the Preferences snapshot fields and debug selectors for viewer/editor action defaults do not exist. GREEN build: `.build/logs/msbuild-20260502_172306_280.log`. GREEN focused run: `cmd_preferences_dialog_viewers_editors_file_action_settings_apply` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_172731` with `preferences.ui.editors_layout_us` values 2,206/137/86/104 us and `preferences.ui.viewers_layout_us` values 6,539/226/205/252 us. Authoritative spec/docs updated in `Specs/Core/Core_SettingsStore.md`, `Specs/UI/UI_CommandMenuKeyboard.md`, `Docs/Preferences.md`, `Docs/Viewers.md`, `Docs/Plugins.md`, `Docs/Todo.md`, and `Docs/UserGuide.md`; `git diff --check` reported no whitespace errors, only existing LF/CRLF normalization warnings.

Phase 3 shell and item utilities:

- [done] `cmd/pane/contextMenuCurrentDirectory`: RED selftest added for active-pane current-folder path routing through a shell-action probe; RED build `.build/logs/msbuild-20260502_173758_346.log`; clean implementation builds `.build/logs/msbuild-20260502_175055_305.log` and `.build/logs/msbuild-20260502_175536_349.log`; GREEN run `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_175302` with `shell.context_menu_current_directory_us` = 13 us; stale `[todo]` guard passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_175712`; authoritative command spec and user guide updated.
- [done] `cmd/pane/openSecurity`: RED selftest added for focused-item Security property-page routing through the same shell-action probe; RED build `.build/logs/msbuild-20260502_173758_346.log`; clean implementation builds `.build/logs/msbuild-20260502_175055_305.log` and `.build/logs/msbuild-20260502_175536_349.log`; GREEN run `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_175257` with `shell.open_security_us` = 26 us; stale `[todo]` guard passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_175712`; authoritative command spec and user guide updated.
- [done] `cmd/pane/changeAttributes` with attribute changes, alternate data stream removal, and operation report.
  - [done] RED coverage for setting/clearing attributes, removing alternate data streams, selected-item scope, and recording an operation report: build failed as expected on missing `ChangeAttributesOptions`, `AttributeChangeState`, report, and debug hooks in `.build/logs/msbuild-20260502_180452_214.log`.
  - [done] Command implementation with localized dialog/options, unsupported-context feedback, refresh, and perf metrics builds cleanly in `.build/logs/msbuild-20260502_182330_300.log`.
  - [done] GREEN focused selftest `cmd_pane_changeAttributes_applies_attributes_removes_streams_and_reports` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_182520` with `fileattrs.stream_enumerate_us` = 336/225 us, `fileattrs.stream_remove_us` = 163/136 us, `fileattrs.feedback_us` = 644 us, and `fileattrs.change_attributes_us` = 78,232 us for two items/two removed streams.
  - [done] Remove stale menu/spec TODO markers and update authoritative command/user docs. GREEN build after resource/doc cleanup: `.build/logs/msbuild-20260502_182841_593.log`; GREEN stale-label guard: `implemented_menu_labels_not_todo` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_183023`.
- [done] `cmd/pane/goToShortcutOrLinkTarget` for `.lnk`, `.url`, reparse points, junctions, and mount points.
  - [done] RED coverage for `.url` local file target navigation and focus restoration: build `.build/logs/msbuild-20260502_183402_378.log`; archived failing run `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_183552` failed with `Go to shortcut/link target should navigate to the target file's parent folder.`
  - [done] Production resolver/dispatch for `.url` local file targets, including localized unsupported/failure feedback and `shell.go_to_shortcut_target_us` perf metric: build `.build/logs/msbuild-20260502_184331_713.log`; GREEN focused run `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_184515` with `shell.go_to_shortcut_target_us` = 14,577 us.
  - [done] RED coverage for `.lnk` file and directory target navigation: build `.build/logs/msbuild-20260502_184740_146.log`; archived failing runs `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_184925` and `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_184945`.
  - [done] Production ShellLink resolver for `.lnk` file and directory targets with COM RAII and no UI prompts: build `.build/logs/msbuild-20260502_185058_072.log`; GREEN runs `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_185238` (`shell.go_to_shortcut_target_us` = 78,972 us) and `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_185254` (`shell.go_to_shortcut_target_us` = 18,833 us).
  - [done] Broken-link and non-local URL policy coverage: build `.build/logs/msbuild-20260502_185445_679.log`; GREEN runs `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_185625` (`shell.go_to_shortcut_target_us` = 2,708 us, `shell.feedback_us` = 990 us) and `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_185642` (`shell.go_to_shortcut_target_us` = 1,206 us, `shell.feedback_us` = 590 us).
  - [done] RED coverage for junction/mount-point reparse target navigation: build `.build/logs/msbuild-20260502_185856_262.log`; archived failing run `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_190046`.
  - [done] Production reparse resolver for mount-point/junction targets with deterministic local navigation: build `.build/logs/msbuild-20260502_190251_220.log`; GREEN run `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_190428` with `shell.go_to_shortcut_target_us` = 20,692 us.
- [done] File properties extension to display shortcut, link, reparse-point, and mount-point target information.
  - [done] RED coverage for `.lnk`, `.url`, and junction/mount-point target fields in item properties: build `.build/logs/msbuild-20260502_190847_635.log`; archived failing run `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_191024` failed on missing `.lnk` `Shortcut` section.
  - [done] Built-in file-system item-properties provider emits Shortcut, Internet Shortcut, and Reparse Point sections with copied yyjson strings. GREEN build `.build/logs/msbuild-20260502_191429_088.log`; GREEN focused run `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_191608` with `itemprops.link_target_us` = 2,241 us for `.lnk`, 146 us for `.url`, and 41 us for reparse metadata.
- [done] `cmd/pane/goToShortcutOrLinkTarget` docs/spec cleanup, menu TODO removal, and stale-marker guards. GREEN rebuild `.build/logs/msbuild-20260502_191948_804.log`; GREEN stale-label guard `implemented_menu_labels_not_todo` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_192128`. Authoritative behavior is now in `Specs/UI/UI_CommandMenuKeyboard.md`, `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, and user docs.
- [done] `cmd/pane/newFromShellTemplate` and `cmd/pane/newFromShellTemplate/<templateId>`.
  - [done] RED coverage for synthetic ShellNew `NullFile`, `Data`, and `FileName` templates, dynamic New-menu population, parameterized template dispatch, created-item focus, and stale-id feedback: build fails on missing ShellNew template model/debug provider/prompt answer hooks in `.build/logs/msbuild-20260502_192856_063.log`.
  - [done] Production ShellNew template enumeration, creation, dynamic menu mapping, parameterized dispatch, prompt integration, and perf metrics. Clean rebuild `.build/logs/msbuild-20260502_194524_446.log`; GREEN focused run `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_194710` passed `cmd_pane_newFromShellTemplate_creates_null_data_and_filename_templates` and `cmd_pane_newFromShellTemplate_menu_and_missing_template_feedback` with `shellnew.create_us` = 792/1,633/1,835/1,003 us, `shellnew.menu_populate_us` = 77 us for two templates, and `shellnew.feedback_us` = 807 us for stale-id feedback.
  - [done] ShellNew docs/spec cleanup, menu TODO removal, and stale-marker guards: `Specs/UI/UI_CommandMenuKeyboard.md`, `Docs/UserGuide.md`, `Docs/FileOperations.md`, and `Docs/Todo.md` now describe the implemented safe ShellNew path and remove stale planned/TODO labels. GREEN stale-label guard `implemented_menu_labels_not_todo` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_195051`.

Phase 4 clipboard, command-line, quick-search, and user menu:

- [done] `cmd/pane/clipboardCut` for file selections while preserving navigation edit cut behavior.
  - [done] RED coverage for local file selection `CF_HDROP` payloads and Preferred DropEffect = move: `cmd_pane_clipboardCut_sets_move_drop_effect` failed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_200214` because the Cut command wrote no `CF_HDROP` paths.
  - [done] Production command routing and file-drop clipboard payload implementation. GREEN build `.build/logs/msbuild-20260502_200816_493.log`; GREEN focused run `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_201002` passed with `clipboard.cut_us` = 935 us for two paths. Fresh post-doc build `.build/logs/msbuild-20260502_201620_298.log`; fresh focused run `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_201840` passed with `clipboard.cut_us` = 992 us for two paths.
- [done] `cmd/pane/clipboardPasteShortcut` for creating `.lnk` files while preserving navigation edit paste behavior.
  - [done] RED coverage for clipboard `CF_HDROP` sources, unique `.lnk` names, refresh, target metadata, and focus: `cmd_pane_clipboardPasteShortcut_creates_unique_links` failed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_200230` because no shortcut files appeared.
  - [done] Production shortcut creation from clipboard file drops. GREEN build `.build/logs/msbuild-20260502_200816_493.log`; GREEN focused run `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_201009` passed with `clipboard.paste_shortcut_us` = 43,174 us for two created links. Fresh post-doc focused run `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_201847` passed with `clipboard.paste_shortcut_us` = 39,583 us for two created links.
  - [done] Empty/non-file clipboard negative path shows localized pane feedback and creates no shortcuts: `cmd_pane_clipboardPasteShortcut_rejects_missing_clipboard_paths` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_201853` with `clipboard.feedback_us` = 12,711 us.
  - [done] Docs/spec/menu TODO cleanup and guard coverage: `Specs/UI/UI_CommandMenuKeyboard.md`, `Docs/FileOperations.md`, `Docs/UserGuide.md`, `Docs/Todo.md`, English resources, and French resources now describe Cut/Paste Shortcut as implemented. GREEN build `.build/logs/msbuild-20260502_201620_298.log`; `implemented_menu_labels_not_todo` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_201920`; `cmd_pane_navigation_change_directory_edit_clipboard_accelerators` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_201927`.
- [done] `cmd/pane/quickSearch` integrated pane search: highlight all matches, select starts-with best match, keep keyboard navigation inside items, Escape/Enter behavior.
  - [done] Audit existing `FolderView` incremental-search engine, command routing, and selftest hooks before writing the RED command-activation coverage.
  - [done] RED selftest coverage for command activation, all-match highlighting, starts-with selection, match navigation, no-match state, Enter accept, and Escape clear. RED build: `.build/logs/msbuild-20260502_202658_710.log` failed on missing incremental-search debug snapshot hooks.
  - [done] Production implementation: Quick Search command routing, pane activation API, match navigation across all matches, Enter accept semantics, debug snapshot, and quick-search perf counters. GREEN build: `.build/logs/msbuild-20260502_203140_449.log`; GREEN focused run: `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_203319` passed with `quicksearch.activate_us` = 676/306 us, `quicksearch.update_us` = 380/69/101 us, `quicksearch.navigate_us` = 690/209 us, and `render.incremental_search_effect_updates` evidence.
  - [done] Docs/spec/menu cleanup for implemented Quick Search while keeping command-line input as a separate follow-up slice. Fresh GREEN build: `.build/logs/msbuild-20260502_203609_175.log`; fresh GREEN Quick Search run: `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_203752`; stale-label guard: `implemented_menu_labels_not_todo` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_203819`.
- [done] `cmd/pane/bringCurrentDirToCommandLine`.
- [done] `cmd/pane/bringFilenameToCommandLine`.
  - [done] RED selftest coverage for command-line host visibility/focus, quoted current-directory insertion, focused filename append, selected path append, Enter launch, and working-directory propagation.
  - [done] RED build captured: `.build/logs/msbuild-20260502_204228_710.log` failed on missing command-line host debug APIs.
  - [done] Existing-host audit: no separate command-line UI/host exists; the navigation address edit is intentionally separate from this feature. Registered command ids and default shortcuts exist, but WM-command routing currently falls through to the not-implemented message.
  - [done] Behavior clarification for this implementation slice: insertion commands append at the command-line caret with one separating space when needed; a focused item inserts its display name, while an explicit multi-selection inserts full local paths with the focused item first and the remaining selected items in pane order.
  - [done] Production implementation: command-line child input, layout/focus, insertion quoting, focused/selected item rules, Enter execution via `ShellExecuteExW`, and perf counters.
  - [done] Production build fix: `.build/logs/msbuild-20260502_205639_390.log` failed on command-line layout `LONG`/`int` min/max ambiguity and one local-name shadowing warning; clean rebuild passed in `.build/logs/msbuild-20260502_205838_120.log`.
  - [done] Focused GREEN run: `cmd_pane_command_line_insertion_and_execute` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_210021` with `commandline.insert_current_dir_us` = 11,082 us, `commandline.insert_filename_us` = 1,447/1,272 us, `commandline.focus_to_visible_us` = 4,749/135/69 us, and `commandline.launch_us` = 3,848 us through the selftest launch callback.
  - [done] Docs/spec/menu cleanup for implemented command-line input commands: `Specs/UI/UI_CommandMenuKeyboard.md`, `Docs/UserGuide.md`, `Docs/MainWindow.md`, `Docs/Todo.md`, English resources, and French resources now document the command-line host and have no stale `[todo]` labels. Fresh rebuild passed in `.build/logs/msbuild-20260502_210331_876.log`; fresh command-line run passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_210509` with `commandline.insert_current_dir_us` = 8,454 us, `commandline.insert_filename_us` = 1,221/1,174 us, `commandline.focus_to_visible_us` = 4,633/121/91 us, and `commandline.launch_us` = 3,486 us; stale-label guard passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_210515`.
- [done] User Menu settings model, schema, preferences page, dynamic menu, and `cmd/pane/userMenu` picker command.
  - [done] RED/GREEN selftest coverage for settings/schema roundtrip, Preferences User Menu page visibility, dynamic menu filtering/enabled state, and external launch is in place.
  - [done] RED build captured: `.build/logs/msbuild-20260502_211800_181.log` fails on missing `Common::Settings::Settings::userMenu`, User Menu preference resources/snapshot fields, and command feedback resources.
  - [done] Settings/schema GREEN: `settings_store_view_edit_actions_roundtrip` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_214020`, including the new `userMenu` settings bucket and schema v12 persistence.
  - [done] GREEN implementation patched: settings/schema, localization resources, Preferences User Menu page, preference snapshot plumbing, runtime menu population, generated WM ids, direct parameterized dispatch, and external launch routing are in place.
  - [done] Build fix verified: `.build/logs/msbuild-20260502_213442_185.log` failed in `Preferences.UserMenu.cpp` because the new page used a non-existent DxUi font role; the page now uses the existing Preferences header role and the rebuild passed in `.build/logs/msbuild-20260502_213835_954.log`.
  - [done] Dynamic menu failure investigation: `user_menu_populates_and_dispatches_configured_actions` failed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_214042` because the loaded User Menu did not contain the applicable enabled action plus the disabled missing-executable action.
  - [done] Dynamic menu root cause narrowed: diagnostic run `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_214442` confirms `CollectUserMenuItems(Left)` returns the expected two configured entries before popup rebuild, so the remaining defect was in cached popup handle discovery/rebuild binding.
  - [done] Popup-handle discovery fix patched: User Menu now resolves from the full main-menu tree instead of depending on the cached Files submenu; rebuild passed in `.build/logs/msbuild-20260502_214821_534.log`.
  - [done] Dynamic menu GREEN: `user_menu_populates_and_dispatches_configured_actions` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_215009` with `usermenu.menu_populate_us` = 71 us for two entries, `usermenu.launch_us` = 31,881 us, selected-paths file write/cleanup coverage, and disabled missing-executable state verification.
  - [done] Preferences User Menu page GREEN: `cmd_preferences_dialog_viewers_editors_file_action_settings_apply` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_215119` with `preferences.ui.usermenu_layout_us` = 1,370 us.
  - [done] Docs/spec/menu cleanup: English/French User Menu `[todo]` labels are removed; `Docs/Todo.md`, `Docs/UserGuide.md`, `Docs/Preferences.md`, `Docs/SettingsFile.md`, `Specs/UI/UI_CommandMenuKeyboard.md`, and `Specs/Core/Core_SettingsStore.md` now describe the implemented User Menu behavior and settings contract. Fresh rebuild passed in `.build/logs/msbuild-20260502_215433_200.log`; stale-label guard `implemented_menu_labels_not_todo` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_215619`.
- [done] `cmd/pane/userMenu/<itemId>` parameterized shortcut/menu routing with external-command macros, selected-paths-file lifecycle, invalid-id feedback, and disabled/unavailable reporting.
  - [done] Direct parameterized/unavailable GREEN: `user_menu_parameterized_command_reports_unavailable` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_215052` with `usermenu.feedback_us` = 703 us and `usermenu.launch_us` = 1,339 us for the disabled item.

Phase 5 pane view options:

- [done] Phase 5 live slice A RED selftests patched: `settings_store_pane_view_options_roundtrip` and `pane_view_options_toggle_file_extensions_navigation_filter_bar` cover settings v13 persistence, display-only extension hiding, active/explicit navigation-bar toggles, filter-bar visibility, preserved filter state, and toggle latency artifact output.
- [done] Phase 5 live slice A RED build captured: `.build/logs/msbuild-20260502_220554_275.log` fails on missing pane view-option settings fields, `FolderWindow` toggle/debug APIs, and command routing, proving the new tests are guarding unimplemented behavior.
- [done] Phase 5 live slice A GREEN: build `.build/logs/msbuild-20260502_223412_840.log`; focused runs `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_223600` (`settings_store_pane_view_options_roundtrip`), `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_223555` (`pane_view_options_toggle_file_extensions_navigation_filter_bar`), and `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_223607` (`implemented_menu_labels_not_todo`) passed after the filter-bar sync regression from `2026-05-02_222959` was fixed. Perf evidence: itemCount 3; `paneViewOptions.fileExtensionsToggleUs` 8,031 us, `paneViewOptions.navigationToggleUs` 23,167 us, `paneViewOptions.filterToggleUs` 32,414 us.
- [done] Phase 5 live slice A docs/spec closeout: `Specs/UI/UI_CommandMenuKeyboard.md`, `Specs/UI/UI_FolderWindow.md`, `Specs/Core/Core_SettingsStore.md`, `Docs/UserGuide.md`, `Docs/MainWindow.md`, `Docs/NavigationAndPaths.md`, `Docs/SettingsFile.md`, and `Docs/Todo.md` describe the implemented file-extension/filter/navigation/status behavior. Fresh build `.build/logs/msbuild-20260502_224321_563.log`; stale-label guard `implemented_menu_labels_not_todo` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_224457`.
- [done] `cmd/pane/viewOptions/toggleFileExtensions`.
- [done] `cmd/pane/viewOptions/toggleThumbnails` RED selftests patched: settings round-trip now expects schema v14 `view.thumbnailsVisible`, and `pane_view_options_toggle_thumbnails` requires active-pane shortcut routing, larger DPI-aware thumbnail target, bounded queued work, deterministic fallback completion, pending-work cancellation on disable, and archived perf artifact output.
- [done] `cmd/pane/viewOptions/toggleThumbnails` RED build captured: `.build/logs/msbuild-20260502_225021_231.log` fails on missing `FolderViewSettings::thumbnailsVisible`, `FolderWindow::SetThumbnailsVisible`, thumbnail debug provider mode, and thumbnail debug metrics.
- [done] `cmd/pane/viewOptions/toggleThumbnails` GREEN implementation: settings/schema, pane state/routing, async thumbnail queue, shell thumbnail extraction with fallback, deterministic selftest provider, and command routing build and pass focused selftests/perf evidence.
  - [done] Settings/schema v14 skeleton is patched for `view.thumbnailsVisible`.
  - [done] FolderView thumbnail mode, bounded async queue, UI-thread bitmap creation, cancellation, and debug counters build cleanly.
  - [done] FolderWindow wrappers/debug snapshot and command/menu routing build cleanly.
  - [done] Settings/schema GREEN run: `settings_store_pane_view_options_roundtrip` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_231226`, proving schema v14 and `view.thumbnailsVisible` persistence.
  - [done] Thumbnail UI/perf GREEN run: `pane_view_options_toggle_thumbnails` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_231310` with `paneViewOptions.thumbnailsToggleUs` = 46,615 us, `thumbnails.queued` = 18, `thumbnails.completed` = 18, `thumbnails.fallback` = 18, and `thumbnails.staleDrops` = 0.
  - [done] Stale-label guard and authoritative docs/spec cleanup for thumbnails are patched and verified: English/French menus no longer carry thumbnail `[todo]`; `UI_CommandMenuKeyboard`, `UI_FolderWindow`, `UI_FolderView`, `Core_SettingsStore`, and user docs now describe thumbnail behavior and schema v14. Rebuild after cleanup passed in `.build/logs/msbuild-20260502_231836_066.log`; stale-label guard `implemented_menu_labels_not_todo` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_232031`.
  - [done] Fresh post-cleanup thumbnail behavior/perf rerun passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_232106` with `paneViewOptions.thumbnailsToggleUs` = 46,167 us, `thumbnails.queued` = 18, `thumbnails.completed` = 18, `thumbnails.fallback` = 18, and `thumbnails.staleDrops` = 0.
  - [done] Fresh post-cleanup settings/schema v14 rerun passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_232139`.
- [done] `cmd/pane/viewOptions/togglePreviewPane` with opposite-pane preview host, Folder/Preview tabs, selection-driven reload, navigation/status region integration, and function-bar hidden layout.
  - [done] Audit existing FolderWindow layout, viewer hosting, selection-change hooks, and selftest debug surfaces before writing RED tests.
  - [done] RED selftests are patched for preview toggle, opposite-pane tabs, selection-driven reload, close behavior, and function-bar-hidden layout. RED build captured in `.build/logs/msbuild-20260502_232724_069.log`, failing on missing preview debug/API surface.
  - [done] GREEN implementation for preview host state, tab strip, preview lifecycle, layout, command routing, and bounded local preview loading.
    - [done] Patch `FolderWindow` preview child creation, tab notification handling, selection/focus reload hooks, and command routing.
    - [done] Patch pane layout/show-hide logic so Preview occupies the opposite pane up to the function bar or bottom edge.
    - [done] Patch deterministic text/folder/unsupported preview content and debug snapshot fields.
    - [done] GREEN implementation build passed in `.build/logs/msbuild-20260502_234119_860.log`.
    - [done] Focused preview behavior/perf selftest `pane_view_options_toggle_preview_pane_tabs_and_selection` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_234306` with `paneViewOptions.previewToggleUs` = 45,214 us, `paneViewOptions.previewSwitchTabUs` = 47,942 us, and `preview.bytes` = 17.
    - [done] Function-bar-hidden layout selftest `pane_view_options_preview_pane_extends_without_function_bar` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_234345`.
  - [done] Stale-label/docs/spec closeout for preview pane is patched and verified: English/French menu labels no longer mark preview `[todo]`, `UI_CommandMenuKeyboard`, `UI_FolderWindow`, and user docs now describe opposite-pane Folder/Preview tab behavior.
  - [done] Rebuild after preview docs/resource cleanup passed in `.build/logs/msbuild-20260502_234711_345.log`.
  - [done] Stale-label guard `implemented_menu_labels_not_todo` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_234905`.
- [done] `cmd/pane/viewOptions/toggleFilterBar`.
- [done] `cmd/pane/viewOptions/toggleNavigationBar`.
- [done] Generic shortcut routing for `cmd/pane/viewOptions/toggleStatusBar`; explicit left/right status bar menu commands stay implemented. RED: active-pane toggle did not change the left pane status bar in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_132052`; GREEN: active left/right pane routing passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-02_133030`.

Phase 6 lists, associations, shares, archive operations:

- [done] Make File List settings model and persistence for last selected format, fields, macro text output, clipboard/file target, recursion, and selection/current-folder source. RED selftest patched: `settings_store_make_file_list_roundtrip`; RED build `.build/logs/msbuild-20260503_002510_245.log` fails on the missing settings model and schema fields. C++ model/parser/writer, schema v15, and docs build cleanly in `.build/logs/msbuild-20260503_005611_159.log`; focused GREEN run `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_005737` passed, and `Specs/SettingsStore.schema.json` parses as JSON.
- [done] `cmd/pane/makeFileList` JSON, text, and CSV generation with deterministic output tests and large-list perf evidence. RED selftest patched: `cmd_pane_makeFileList_generates_formats_and_saves_options`; RED build `.build/logs/msbuild-20260503_002510_245.log` fails on the missing command/debug automation and output helper surface. Command routing, resource-backed options dialog, automation hook, generation, file/clipboard output, last-option saving, docs/menu cleanup, and perf counters build cleanly in `.build/logs/msbuild-20260503_005611_159.log`; focused GREEN run `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_005745` passed with JSON artifact `perf/make_file_list_metrics.json` (`entryCount` = 4, test-measured JSON command duration = 95,435 us) and command metrics including `makeFileList.collect_us`, `makeFileList.generate_us`, `makeFileList.output_us`, `makeFileList.feedback_us`, and `makeFileList.total_us` for JSON/CSV/text. Stale-label guard passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_005751`.
- [done] `cmd/pane/listOpenedFiles`: RED selftest patched for empty-state dialog visibility, viewer/preview/external-editor source rows, closed external-editor pruning, row focus navigation, and archived perf artifact evidence. RED build `.build/logs/msbuild-20260503_010701_401.log` fails on missing opened-files dialog/model/debug APIs. Production command/debug declarations, resources, localized strings, process-handle capture, dialog collection/focus behavior, viewer/editor launch tracking, WM-command dispatch, docs/menu/spec/test-coverage cleanup, and stale-label guard coverage are complete. GREEN builds: `.build/logs/msbuild-20260503_012931_250.log` and `.build/logs/msbuild-20260503_013820_113.log`. Focused behavior/perf rerun passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_014045` with `rowCount` = 3 and `listOpenedFiles.open_us` = 93,454 us. Stale-label guard passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_014104`.
- [done] `cmd/pane/shares`: RED selftest is patched for synthetic share rows, sorted display, open-path navigation, access-denied empty/error state, shortcut dispatch, and archived perf artifact output. RED build `.build/logs/msbuild-20260503_014506_541.log` fails on the missing Shared Directories debug provider, snapshot, dialog-control, and command surface.
  - [done] Test-only provider contract and expected debug snapshot behavior are defined by the RED selftest.
  - [done] Production implementation: enumerate local disk shares, show modeless Shared Directories dialog, sort rows, open local share paths, expose Windows Shared Folders management action, and provide deterministic debug hooks. Build passed in `.build/logs/msbuild-20260503_015901_797.log`.
  - [done] Verification: focused GREEN behavior/perf selftest passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_020538`, including `perf/shared_directories_metrics.json`.
  - [done] Closeout: stale menu `[todo]` is removed, stale-label guard entry is added, and command specs/user docs/test coverage are patched; rebuild `.build/logs/msbuild-20260503_020351_523.log`, stale-label guard `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_020528`.
- [done] `cmd/app/rereadAssociations` reloads settings-backed viewer/editor/User Menu/plugin/extension associations, rebuilds dynamic menus, clears normal and association icon caches, refreshes both panes, preserves live pane paths, and keeps prior runtime settings on reload failure. RED focused failure: `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_000742`. GREEN builds: `.build/logs/msbuild-20260503_001148_516.log` and `.build/logs/msbuild-20260503_001755_942.log`. GREEN focused runs: `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_001328` and fresh post-doc/resource run `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_001923` (`rereadAssociations.total_us` = 323,179 us, cache seed count = 50, action counts = 1/1/1). Docs/spec/menu cleanup is verified by stale-label guard `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_001929`.
- [done] `cmd/pane/pack`: local stored-ZIP creation is implemented and verified.
  - [done] Refresh checklist and implementation scope.
  - [done] RED selftest for selected local items packed into deterministic `.zip` output with overwrite validation and perf artifact output. RED build `.build/logs/msbuild-20260503_021445_724.log`.
  - [done] Production command implementation for local file-system selections, settings/debug automation, localized success/failure feedback, and menu/shortcut routing.
    - [done] Stored ZIP writer and pack command plumbing build cleanly in `.build/logs/msbuild-20260503_023247_097.log`.
  - [done] Focused GREEN selftest and archive perf metrics under `Specs/TestRuns/`: `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_023445` reports `archive.pack_us` = 81,016 us, `entryCount` = 16, and `bytesProcessed` = 23.
- [done] `cmd/pane/unpack`: local stored-ZIP extraction is implemented and verified.
  - [done] Refresh checklist and implementation scope.
  - [done] RED selftest for stored ZIP extraction, destination validation, overwrite behavior, unsupported archive/provider feedback, and perf artifact output. RED build `.build/logs/msbuild-20260503_021445_724.log`.
  - [done] Production command implementation for supported archive selections, destination creation, traversal-safe extraction, localized feedback, and refresh behavior.
    - [done] Stored ZIP reader and traversal-safe extractor build cleanly in `.build/logs/msbuild-20260503_023247_097.log`.
  - [done] Focused GREEN selftest and archive perf metrics under `Specs/TestRuns/`: `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_023445` reports `archive.unpack_us` = 74,722 us, `entryCount` = 16, and `bytesProcessed` = 23.

Phase 7 closeout:

- [done] Update `Specs/UI/UI_CommandMenuKeyboard.md`, settings schema spec, user docs, and localization after each completed command slice. Fresh closeout build passed in `.build/logs/msbuild-20260503_024654_942.log`.
- [done] Archive red and green selftest runs plus perf metrics under `Specs/TestRuns/`. Fresh closeout runs passed: archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-03_024835`, registry guard `2026-05-03_024843`, shortcut guard `2026-05-03_024848`, active-pane routing guard `2026-05-03_024856`, and stale-label guard `2026-05-03_024902`.
- [done] Move this plan to `Specs/Plans/Done/` only after all durable behavior has moved into authoritative specs.

> For agentic workers: execute this plan phase by phase. Do not combine high-risk phases unless the previous phase is green. Keep each change buildable, add deterministic command selftests with the implementation, archive required perf evidence, and move durable behavior into authoritative specs before closing.

## Goal

Replace every command that is still presented as planned or routed to the localized "not implemented" alert with concrete behavior, while also cleaning up stale `[td]` / `[todo]` labels for commands that already work.

The work is complete only when:

- Menu labels, command registry metadata, shortcut dispatch, pane routing, docs, specs, and localization agree.
- No command listed in the canonical UI command spec is accidentally left as a silent no-op.
- Commands that remain intentionally unsupported by the active context are disabled or show a precise localized explanation.
- New command behavior has deterministic selftest coverage.
- Perf-sensitive commands have instrumentation and archived evidence under `Specs/TestRuns/`.
- The WIP plan is moved to `Specs/Plans/Done/` after specs and docs are updated.

## Current Evidence

The command backlog was identified from these sources:

- `Docs/Todo.md`: user-facing list of planned UI, command, editor, packing, view, and navigation features.
- `Specs/UI/UI_CommandMenuKeyboard.md`: canonical command ids and `[td]` markers for commands that were planned when the spec was written.
- `RedSalamander/RedSalamander.rc` and `RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.rc`: menu labels containing `[todo]`.
- `RedSalamander/CommandRegistry.cpp`: canonical command ids, WM command ids, and shortcut-visible commands.
- `RedSalamander/RedSalamander.cpp`: `DispatchShortcutCommand`, `OnMainWindowCommand`, dynamic menu maps, and `ShowCommandNotImplementedMessage`.

Important implementation facts:

- `ShowCommandNotImplementedMessage(HWND, std::wstring_view)` already provides the localized fallback.
- `DispatchShortcutCommand(...)` sends WM commands when `TryGetWmCommandId(...)` returns a value; otherwise it shows the fallback.
- `OnMainWindowCommand(...)` shows the fallback from the default branch when a WM command is in the registry but has no explicit handler.
- Several menu labels/spec markers are stale because the command already has a handler.

## Backlog Inventory

### Already implemented commands to protect or clean up

These entries are already implemented. For entries that still have stale `[td]` or `[todo]` markers, remove the markers after focused selftests confirm the behavior. For `cmd/app/viewWidth`, preserve the existing splitter-width behavior and documentation.

- `cmd/pane/permanentDelete`
- `cmd/pane/selection/selectSameName`
- `cmd/pane/selection/unselectSameName`
- `cmd/pane/sort/attributes`
- `cmd/app/plugins/manage`
- `cmd/pane/selectFileSystemPlugin`
- `cmd/app/theme/select`
- `cmd/pane/windowMenu`
- `cmd/app/toggleFunctionBar`
- `cmd/app/toggleMenuBar`
- `cmd/app/viewWidth`, which is already the implemented splitter view-width adjust command and must not be repurposed as the viewer picker
- `cmd/pane/viewOptions/toggleStatusBar`, for menu commands only
- `IDM_RIGHT_HOT_PATHS`

The cleanup task must update:

- `Specs/UI/UI_CommandMenuKeyboard.md`
- `RedSalamander/RedSalamander.rc`
- `RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.rc`
- Any command inventory selftest expectation that still treats these commands as planned

### Real unimplemented command families

These commands still need implementation work or routing fixes:

- Viewer/editor family:
  - `cmd/pane/alternateView`
  - `cmd/pane/viewWith`
  - `cmd/pane/edit`
  - `cmd/pane/editWith`
  - `cmd/pane/alternateEdit`
  - `cmd/pane/editNew`
- Shell and item utilities:
  - `cmd/pane/contextMenuCurrentDirectory`
  - `cmd/pane/openSecurity`
  - `cmd/pane/changeAttributes`
  - `cmd/pane/goToShortcutOrLinkTarget`
  - `cmd/pane/newFromShellTemplate`
- Clipboard and command-line family:
  - `cmd/pane/clipboardCut`, outside navigation edit controls
  - `cmd/pane/clipboardPasteShortcut`
  - `cmd/pane/quickSearch`
  - `cmd/pane/bringCurrentDirToCommandLine`
  - `cmd/pane/bringFilenameToCommandLine`
  - `cmd/pane/userMenu`
- Pane view options:
  - `cmd/pane/viewOptions/toggleFileExtensions`
  - `cmd/pane/viewOptions/toggleThumbnails`
  - `cmd/pane/viewOptions/togglePreviewPane`
  - `cmd/pane/viewOptions/toggleFilterBar`
  - `cmd/pane/viewOptions/toggleNavigationBar`
  - Generic shortcut routing for `cmd/pane/viewOptions/toggleStatusBar`
- Lists, associations, shares, archives:
  - `cmd/pane/makeFileList`
  - `cmd/pane/listOpenedFiles`
  - `cmd/pane/shares`
  - `cmd/app/rereadAssociations`
  - `cmd/pane/pack`
  - `cmd/pane/unpack`
- Parameterized/dynamic command routing gaps and new theme cycle commands:
  - `cmd/app/theme/select/<themeId>`
  - `cmd/app/plugins/toggleEnabled/<pluginId>`
  - `cmd/app/plugins/configure/<pluginId>`
  - `cmd/pane/selectFileSystemPlugin/<pluginId>`
  - `cmd/pane/viewWith/<viewerId>`
  - `cmd/pane/editWith/<editorId>`
  - `cmd/pane/newFromShellTemplate/<templateId>`
  - `cmd/pane/navigatePath/<pathId>`
  - `cmd/app/openFileExplorerKnownFolder/<knownFolderId>`
  - `cmd/app/theme/selectNext`
  - `cmd/app/theme/selectPrev`

## Per-Command Implementation Specification

This section is the command-by-command contract for the missing command backlog. Later phases describe implementation order; this section defines what "implemented" means.

Rules that apply to every command below:

- Menu activation and shortcut activation must execute the same command path.
- A command must never silently do nothing. It must run, be disabled in the menu, or show a localized unsupported-context message when invoked by shortcut.
- All user-visible text must come from resources.
- Commands that operate on pane contents target the active pane from shortcuts and the explicit left/right pane from pane-specific menus.
- Commands that start background work must support cancellation where the existing operation framework supports it.
- Commands that touch shell, COM, registry, Win32 handles, GDI objects, or global memory must use WIL RAII wrappers.
- Commands that can affect responsiveness, rendering, file operations, queueing, or memory retention must add metrics, deterministic selftests, and archived runs under `Specs/TestRuns/`.

Existing implemented command guard:

- `cmd/app/viewWidth` is already implemented as the splitter view-width adjust command with selftests and a function-bar shortcut. It must remain documented as that command in `Specs/UI/UI_CommandMenuKeyboard.md`, command registry metadata, menus, shortcut defaults, and command selftests. The viewer picker command is `cmd/pane/viewWith`; this plan must not repurpose `cmd/app/viewWidth`.
- `cmd/pane/editWidth` is removed from this implementation scope. The implementation plan must remove or retire that command from menus, default shortcuts, command registry metadata, and specs unless a compatibility import path is required for old shortcut settings.
- Add the missing `cmd/pane/alternateEdit` command beside `cmd/pane/edit`.

### Viewer and Editor Commands

| Command | Required behavior | Enablement and invalid context | Required validation |
| --- | --- | --- | --- |
| `cmd/pane/viewWith` | Display a picker of available viewers for the focused file. The picker includes internal viewer plugins and external viewer actions configured in settings. It is filtered by file extension, computer name, enabled state, and action availability. | Enable only when the active pane has a focused item that can be viewed. Disable for unsupported plugin file-system entries. Shortcut invocation without a valid file shows a localized "no file to view" message. | Selftest picker population, extension filtering, computer-name filtering, internal plugin action, external program action, and localized launch failure. |
| `cmd/pane/viewWith/<viewerId>` | Run the exact viewer action selected by id. The id can represent either an internal viewer plugin action or an external viewer action from settings. | Same as `cmd/pane/viewWith`; additionally validate the action id exists, is enabled, applies to the file extension/computer, and can resolve required macros. | Selftest valid internal id, valid external id, missing id, disabled id, computer mismatch, extension mismatch, and paths with spaces. |
| `cmd/pane/alternateView` | Run the configured alternate viewer action for the focused file. The alternate viewer is configured beside the normal `View` association in settings and is not automatically "the next viewer." It can be an internal viewer plugin or an external program using macros. | Enable wherever the focused item has a valid alternate-view action for its extension/computer. If no alternate action is configured, show a localized "no alternate viewer configured" message. | Selftest extension-specific alternate action, computer-specific alternate override, internal plugin alternate, external program alternate, macro expansion, and missing configuration. |
| `cmd/pane/edit` | Run the configured primary editor action for the focused file. The editor action can be an external program with macros and can vary by extension and computer name. | Enable only when the focused item has a valid editor action and can provide a local path or other configured macro inputs. Shortcut invocation without a valid file shows a localized "no file to edit" message. | Selftest default editor resolution, extension-specific editor, computer-specific editor, macro expansion, dry-run launch command, launch failure reporting, and pane refresh request after launch. |
| `cmd/pane/editWith` | Display a picker of available editor actions for the focused file. The picker includes configured editor actions filtered by extension, computer name, enabled state, and executable availability. | Enable only when a focused item can be edited and at least one editor action is available. Missing editor executables are shown disabled with a reason. | Selftest picker population, missing executable disabled state, preferences route, extension/computer filtering, and selected editor dry-run launch. |
| `cmd/pane/editWith/<editorId>` | Run the exact editor action selected by id from shortcut or menu. | Same as `cmd/pane/edit`; additionally validate the editor id exists, is enabled, applies to the extension/computer, has a launchable executable, and can resolve required macros. | Selftest valid id, missing id, disabled id, missing executable, extension mismatch, computer mismatch, and argument token expansion. |
| `cmd/pane/alternateEdit` | Run the configured alternate editor action for the focused file. The alternate editor is configured beside the normal `Edit` association in settings and can be an external program with macros. | Enable when the focused item has a valid alternate-edit action for its extension/computer. If not configured, show a localized "no alternate editor configured" message. | Selftest new command registry entry, menu label, shortcut dispatch, extension-specific alternate edit, computer override, external command macro expansion, and missing configuration. |
| `cmd/pane/editNew` | Prompt for a new file name, infer the file extension, show an `Editor` combo populated from configured editor actions that apply to that extension and this computer, create the file, focus it, and launch the selected editor. | Enable only for writable local directories or plugin file systems with a create-file capability. Invalid names, absolute paths, parent traversal, reserved names, and collisions must be rejected with localized validation. The Editor combo is disabled only when no editor action is available. | Selftest editor combo filtering by extension/computer, invalid names, collision handling, temp-file creation, focus request, selected editor launch dry-run, and refresh behavior. |

### Shell and Item Utility Commands

| Command | Required behavior | Enablement and invalid context | Required validation |
| --- | --- | --- | --- |
| `cmd/pane/contextMenuCurrentDirectory` | Open the Windows shell context menu for the active pane directory itself. Do not target the focused item. Refresh the pane after the menu closes if a shell verb may have changed contents. | Enable for local file-system directories. Plugin file-system panes show localized unsupported-context text when invoked by shortcut. | Selftest target path resolution, local directory route, plugin unsupported route, and refresh request after simulated shell verb. |
| `cmd/pane/openSecurity` | Open the Windows Security property page for the focused item. If no item is focused, target the current directory. | Enable for local files and directories. Disable for unsupported plugin entries. Shortcut invocation on unsupported entries shows localized security-page unsupported text. | Selftest file target, directory target, current-directory fallback, unsupported provider, and shell API failure. |
| `cmd/pane/changeAttributes` | Open an attributes dialog for selected local files/directories. Apply supported attribute changes with `SetFileAttributesW`, optionally remove alternate data streams from the selection, and show an operation report covering attribute changes and stream removals. | Enable for one or more local file-system items. Multi-selection shows mixed state. Stream removal is available only for NTFS/local paths where streams can be enumerated safely. Unsupported plugin file systems are disabled or report localized unsupported-context text. | Selftest single item, multi-item mixed state, read-only/archive toggles on temp files, alternate data stream enumeration/removal on temp files, partial failure report, and refresh request. |
| `cmd/pane/goToShortcutOrLinkTarget` | Resolve the focused `.lnk`, `.url`, reparse point, junction, or mount point target. Navigate to the target directory and focus the target file when local; navigate directly when the target is a directory. Also extend file properties to display link/reparse/mount-point target information. | Enable for focused shortcut/link/reparse-point entries. Broken links, inaccessible targets, and unsupported targets show localized messages. URL targets follow the documented URL-opening policy. | Selftest `.lnk` to file, `.lnk` to directory, broken `.lnk`, `.url` local path, `.url` web URL policy, junction/reparse target, mount point target, file-properties target display, and focus behavior. |
| `cmd/pane/newFromShellTemplate` | Show a dynamic submenu of ShellNew templates and create a new file from the selected template in the active directory. Support `NullFile`, `FileName`, `Data`, and safe `Command` templates. | Enable only for writable local directories. Empty template list disables the submenu and can show a localized "no templates available" message from shortcut. | Selftest synthetic template enumeration, invalid file name rejection, unique name creation, each supported template form, and created-item focus. |
| `cmd/pane/newFromShellTemplate/<templateId>` | Create a new file from the exact shell template id from the command parameter. | Same as `cmd/pane/newFromShellTemplate`; additionally validate that the template id still exists. Missing id shows localized template unavailable text. | Selftest valid id, missing id, stale id after reread associations, and path with spaces. |

### Clipboard, Command-Line, Quick Search, and User Menu Commands

| Command | Required behavior | Enablement and invalid context | Required validation |
| --- | --- | --- | --- |
| `cmd/pane/clipboardCut` | When focus is in a navigation edit control, preserve text cut behavior. Otherwise, place selected local files on the clipboard as `CF_HDROP` with Preferred DropEffect `DROPEFFECT_MOVE`. | Enable for editable text focus or one or more local file-system selections. Disable for unsupported plugin entries. Shortcut on unsupported entries shows localized clipboard unsupported text. | Selftest text edit cut preservation, file selection clipboard payload, move drop effect, unsupported provider, and empty selection. |
| `cmd/pane/clipboardPasteShortcut` | When focus is in a navigation edit control, preserve text paste behavior if that control owns the command. Otherwise, create `.lnk` shortcuts in the active directory for paths currently on the clipboard. | Enable for writable local directories when clipboard contains local file paths. Invalid clipboard contents show localized "no shortcut source" text from shortcut. | Selftest edit-control path, one source path, multiple source paths, unique names, invalid clipboard, and refresh/focus behavior. |
| `cmd/pane/quickSearch` | Activate the pane's integrated quick-search behavior, not the persistent filter bar. Typed characters highlight all matching items, select the best item whose name starts with the search text, and keep keyboard navigation inside the folder items. | Enable when the active pane has an enumerable item list. Shortcut with an unavailable pane shows localized unavailable text. | Selftest integrated search activation, all-match highlighting, starts-with item selection, next/previous keyboard navigation through matches, Escape behavior, Enter behavior, and no-match text. |
| `cmd/pane/bringCurrentDirToCommandLine` | Focus command-line input and insert the active pane directory path using command-line quoting rules. | Enable when the active pane has a current directory path. Unsupported plugin directories show localized text unless they can provide a shell path. | Selftest local path insertion, path with spaces, existing input append/replace rule, and unsupported provider. |
| `cmd/pane/bringFilenameToCommandLine` | Focus command-line input and insert the focused item name or selected item path according to the UI spec. Use command-line quoting rules. | Enable when a file or directory is focused or selected. Empty panes show localized no-item text from shortcut. | Selftest focused item, single selected item, multi-selection rule, path with spaces, and empty pane. |
| `cmd/pane/userMenu` | Open a dynamic menu of configured user commands. Entries launch external commands using the same macro engine as View, Alternate View, Edit, and Alternate Edit. Empty menu opens the User Menu settings page with a localized explanation. | Enable when settings can be read. Individual entries are disabled when their executable is missing, their computer-name rule does not match, or required macros cannot be resolved. | Selftest empty menu route, settings-page route, populated menu, disabled missing executable, computer filtering, shared macro expansion, and dynamic id cleanup. |
| `cmd/pane/userMenu/<itemId>` | Execute the exact configured user menu entry from the command parameter. Expand supported macros and launch through the shared external-action helper. | Validate item id, enabled state, executable, computer-name rule, working directory, and macro requirements. Missing or invalid entries show localized text. | Selftest valid item dry-run, missing id, disabled item, macro expansion, selected-paths temp file creation, and cleanup. |

### Pane View Option Commands

| Command | Required behavior | Enablement and invalid context | Required validation |
| --- | --- | --- | --- |
| `cmd/pane/viewOptions/toggleFileExtensions` | Toggle whether file extensions are shown in the target pane. This is display-only; operations keep using real file names and paths. | Enable for file-system panes with a visible item list. Unsupported panes show localized text from shortcut. | Selftest display label with and without extension, folder names unchanged, operation path unchanged, setting persistence, and menu check state. |
| `cmd/pane/viewOptions/toggleThumbnails` | Toggle thumbnail mode in the target pane. Thumbnails load asynchronously with icon fallback and bounded visible-work scheduling. | Enable for file-system panes. When thumbnail providers are unavailable, keep command enabled but render icon fallback and log no normal-control-flow warning. | Selftest setting persistence, async request scheduling, stale request cancellation, icon fallback, DPI-aware size, and menu check state. |
| `cmd/pane/viewOptions/togglePreviewPane` | Toggle preview mode in the other pane. When preview is open, the other pane shows top tabs for `Folder` and `Preview`; selecting items in the main pane updates the preview tab. Closing preview removes the tab. The preview area includes that pane's navigation bar and status bar region and extends from under the main menu to the function bar, or to the bottom window edge when the function bar is hidden. | Enable when the opposite pane can host preview mode. Unsupported focused files show localized unsupported preview inside the Preview tab, not a modal error. | Selftest opposite-pane host selection, tab creation/removal, selection-driven preview reload, navigation/status bar layout inside preview host, function-bar hidden layout, unsupported file message, cancellation, pane destruction, and setting persistence. |
| `cmd/pane/viewOptions/toggleFilterBar` | Toggle a visible filter bar for the target pane. The bar reflects the current filter state and reuses existing filtering semantics. | Enable for enumerable panes. Shortcut on unavailable panes shows localized text. | Selftest visibility toggle, current filter displayed, closing behavior, focus return, setting persistence, and menu check state. |
| `cmd/pane/viewOptions/toggleNavigationBar` | Toggle the navigation/address bar for the target pane. Address-bar focus commands must show the bar first or report localized unavailable text. | Enable for normal folder panes. Keep left/right explicit menu commands and active-pane shortcut behavior consistent. | Selftest left pane, right pane, active pane, focus-address-bar interaction, setting persistence, and menu check state. |
| `cmd/pane/viewOptions/toggleStatusBar` | Route generic shortcut command to the active pane using the same code as existing left/right status bar menu commands. | Enable for normal folder panes. Unsupported pane shortcut shows localized text. | Selftest generic active-pane route, explicit left/right route, setting persistence, and menu check state. |

### List, Association, Share, Pack, and Unpack Commands

| Command | Required behavior | Enablement and invalid context | Required validation |
| --- | --- | --- | --- |
| `cmd/pane/makeFileList` | Generate a file list from selected items or the current directory. Support JSON, text, and CSV output. Let the user choose which fields are included, define text output through macros, choose clipboard/file output, and persist the last selected options in settings. | Enable for enumerable panes. Plugin file systems are supported only when enumeration can provide required fields; otherwise show localized unsupported-context text. | Selftest deterministic temp folder output, JSON schema/output, CSV escaping, text macro expansion, saved option reload, clipboard output, file output, cancellation, and unsupported provider. |
| `cmd/pane/listOpenedFiles` | Show a dialog listing files opened by RedSalamander viewer instances, launched editor processes tracked by the app, and active preview handles if any. | Always enabled. Empty state shows localized "no opened files" text. | Selftest empty state, viewer instance entry, editor dry-run tracked entry, focus owning pane action, and closed-instance cleanup. |
| `cmd/pane/shares` | Show local shared directories using `NetShareEnum`, including share name, local path, type, and remark where available. | Enable on Windows. Access denied shows localized non-fatal text and offers the management action if available. | Selftest synthetic provider success, empty list, access denied, display fields, and management action route. |
| `cmd/app/rereadAssociations` | Refresh extension associations, viewer/editor dynamic menus, file-system plugin selection menus, icon/thumbnail association caches, and visible panes. | Always enabled. If a provider reload fails, report localized partial-failure text and keep previous valid state where possible. | Selftest settings reload, menu rebuild, stale dynamic ids invalidated, active plugin preserved when still valid, and cache refresh request. |
| `cmd/pane/pack` | Create an archive from selected items or the active directory. Minimum shippable format is `.zip`; add `.7z` if write/update support is added to the existing 7-Zip integration. | Enable for local readable items. Unsupported plugin sources show localized text unless plugin extraction/read APIs can provide stream input safely. | Selftest archive path validation, deterministic archive entries, overwrite policy, many-small-files scenario, one-large-file scenario, cancellation, and progress reporting. |
| `cmd/pane/unpack` | Extract selected archives to a destination folder using the file-operation progress/cancel infrastructure. Preserve directory structure and overwrite policy. | Enable when selected items are archives supported by the existing archive reader. Unsupported formats show localized text. | Selftest small archive extraction, overwrite policy, destination validation, cancellation cleanup, refresh behavior, and progress reporting. |

### Dynamic and Theme Command Routing

| Command pattern | Required behavior | Required validation |
| --- | --- | --- |
| `cmd/app/theme/select/<themeId>` | Select the exact built-in or custom theme id through the same path used by theme menu items. | Selftest built-in id, custom id, missing id, and menu check state refresh. |
| `cmd/app/theme/selectNext` | Cycle to the next available theme in the same order shown by the main theme menu. Wrap from the last theme to the first. Add menu and function-bar shortcut integration. | Selftest built-in/custom ordering, wraparound, menu route, function-bar shortcut route, and menu check state refresh. |
| `cmd/app/theme/selectPrev` | Cycle to the previous available theme in the same order shown by the main theme menu. Wrap from the first theme to the last. Add menu and function-bar shortcut integration. | Selftest built-in/custom ordering, wraparound, menu route, function-bar shortcut route, and menu check state refresh. |
| `cmd/app/plugins/toggleEnabled/<pluginId>` | Toggle the exact plugin id through the plugin manager model. | Selftest enabled-to-disabled, disabled-to-enabled, missing id, and menu rebuild. |
| `cmd/app/plugins/configure/<pluginId>` | Open configuration for the exact plugin id when the plugin exposes configuration UI. | Selftest configurable plugin, non-configurable plugin localized message, and missing id. |
| `cmd/pane/selectFileSystemPlugin/<pluginId>` | Select the exact file-system plugin id for the active pane. | Selftest valid id, missing id, disabled plugin, active-pane targeting, and menu check state. |
| `cmd/pane/navigatePath/<pathId>` | Navigate to the exact path/history/hot-path target represented by the parameter. | Selftest existing target, stale target, path with spaces, and active-pane targeting. |
| `cmd/app/openFileExplorerKnownFolder/<knownFolderId>` | Open the exact known folder in File Explorer through the same helper as menu commands. | Selftest valid known folder, missing id, shell failure, and no command fallback regression. |

## Feature Behavior and Low-Fidelity UX

This section describes what the user experiences when each missing command is implemented. These sketches are low-fidelity behavior diagrams, not final visual design.

Shared interaction rules:

- Menu entries that cannot run in the current context are disabled before the user clicks them.
- Shortcut invocation of an unavailable command shows a localized modeless host alert.
- Commands that affect the active pane use the focused pane unless the menu item explicitly says Left or Right.
- Dialogs remember their last useful choices when that improves repeated use and does not create hidden destructive behavior.
- Long operations show progress, cancellation, and completion/failure state through the existing file-operation UI.

### `cmd/pane/viewWith`

User-facing feature: pick one available viewer action for the focused file.

Flow:

1. User focuses a file.
2. User opens `Files > View With`.
3. RedSalamander builds a viewer list from internal viewer plugins and external viewer actions configured in settings.
4. RedSalamander filters the list by file extension, computer name, enabled state, and launch availability.
5. User chooses a viewer.
6. The file opens through that viewer action. The pane selection and focus stay unchanged.

Low-fi menu:

```text
Files
  View
  View With >
      Text Viewer
      Space Viewer
      Hex Viewer
      External: VS Code Preview
      External: Custom Binary Viewer (disabled: this computer)
      ----------------
      Viewer preferences...
```

Behavior details:

- The normal `View` command keeps using the configured primary view action.
- `View With` is for an explicit one-time choice.
- Viewer actions can be internal plugin viewers or external programs.
- External viewer actions use the shared macro engine:
  - `{path}`: active or parent directory path
  - `{fullPath}`: full item path including the file name
  - `{pathAndFilename}`: explicit alias for full item path
  - `{filename}`: leaf file name only
  - `{selectedPathsFile}`: path to a temporary file containing selected paths
  - `{computer}`: current computer name
- The submenu is rebuilt when associations or plugins are reread.
- If the selected viewer fails to open the file, the user gets a localized message naming the file and viewer.

### `cmd/pane/viewWith/<viewerId>`

User-facing feature: run the exact same "View With this viewer action" behavior from a shortcut or dynamic menu item.

Flow:

1. Shortcut dispatch resolves the viewer action id.
2. RedSalamander validates that the action still exists, is enabled, applies to this extension, and applies to this computer.
3. The focused file opens with that viewer action.

No custom UI is shown unless the viewer action id is invalid or the viewer launch fails.

### `cmd/pane/alternateView`

User-facing feature: open the focused file with the alternate viewer action configured for that file type and computer.

Flow:

1. User focuses `example.log`.
2. User invokes `Alternate View`.
3. RedSalamander resolves the alternate-view action for `.log`, considering computer-specific overrides.
4. RedSalamander runs that action, either through an internal viewer plugin or an external program.

Low-fi behavior:

```text
Focused file: example.log
Settings match:
  extension: .log
  computer: BUILD-BOX-01
  alternate view: External "Log Analyzer"

Expanded command:
  C:\Tools\LogAnalyzer.exe "C:\Work\example.log"
```

Behavior details:

- This is a settings-driven action, not automatic "next viewer" cycling.
- The alternate viewer can be an internal viewer plugin or an external program.
- External actions use the same macro engine as `View With`, `Edit`, `Alternate Edit`, and User Menu.
- If no alternate viewer is configured for the file/computer, show "No alternate viewer is configured for example.log."

### `cmd/pane/edit`

User-facing feature: open the focused file with the primary editor action configured for that file type and computer.

Flow:

1. User focuses a file.
2. User invokes `Edit`.
3. RedSalamander resolves the primary edit action from settings.
4. RedSalamander expands macros such as `{path}`, `{fullPath}`, `{pathAndFilename}`, and `{filename}`.
5. The editor starts with the file.
6. RedSalamander requests a pane refresh after launch.

Low-fi failure:

```text
+------------------------------------------------+
| Editor unavailable                             |
|                                                |
| RedSalamander could not start "Visual Studio". |
| C:\Tools\missing-editor.exe was not found.     |
|                                                |
| [Open Editors Preferences] [Close]             |
+------------------------------------------------+
```

Behavior details:

- The default editor can remain Notepad, but the command is defined by settings and can vary by extension and computer name.
- Paths with spaces are quoted correctly.
- External editor actions share the same macro system as viewer and user-menu actions.
- Editor launch does not block the UI thread.

### `cmd/pane/editWith`

User-facing feature: pick one available editor action for the currently focused file.

Flow:

1. User focuses a file.
2. User opens `Files > Edit With`.
3. RedSalamander shows configured editor actions filtered by extension, computer name, and availability.
4. User chooses an editor.
5. The editor starts with the focused file.

Low-fi menu:

```text
Files
  Edit
  Edit With >
      Notepad
      Visual Studio Code
      Alternate Diff Editor
      Custom Diff Tool     (disabled: missing executable)
      Laptop-only Editor   (disabled: this computer)
      ----------------
      Editors preferences...
```

Behavior details:

- Disabled editor entries remain visible so users can understand why an expected editor is not available.
- The list includes both primary and alternate editor actions when they are valid for the focused file.
- Choosing `Editors preferences...` opens the preferences page focused on editor configuration.

### `cmd/pane/editWith/<editorId>`

User-facing feature: run an exact editor action from shortcut or dynamic menu.

Flow:

1. Shortcut dispatch resolves the editor id.
2. RedSalamander validates that the editor action exists, is enabled, applies to the file extension, applies to this computer, and can launch.
3. The focused file opens with that editor action.

No custom UI is shown unless validation fails.

### `cmd/pane/alternateEdit`

User-facing feature: open the focused file with the alternate editor action configured for that file type and computer.

Flow:

1. User focuses a file.
2. User invokes `Alternate Edit`.
3. RedSalamander resolves the alternate-edit action from settings.
4. RedSalamander expands macros and launches the configured editor.

Low-fi behavior:

```text
Focused file: report.md

Settings match:
  extension: .md
  primary edit: Visual Studio Code
  alternate edit: External "Markdown Preview Editor"

Expanded command:
  C:\Tools\MdPreview.exe "C:\Work\report.md"
```

Behavior details:

- This command must be added to the command registry, menus, resources, shortcut lists, and selftests.
- If no alternate editor is configured, show "No alternate editor is configured for report.md."
- It uses the same external-action macro engine as primary Edit and User Menu.

### `cmd/pane/editNew`

User-facing feature: create a new file in the active directory and immediately edit it.

Flow:

1. User invokes `Edit New`.
2. RedSalamander asks for a file name.
3. RedSalamander infers the file extension as the user types.
4. The `Editor` combo updates to show editor actions that match that extension and this computer.
5. User confirms.
6. RedSalamander creates the file, refreshes the pane, focuses it, and opens it in the selected editor action.

Low-fi dialog:

```text
+-------------------------------------------+
| New file                                  |
|                                           |
| Folder: C:\Work\Notes                     |
| Name:   [ meeting-notes.txt             ] |
| Editor: [ Notepad                    v ]  |
|                                           |
| [Create and Edit] [Cancel]                |
+-------------------------------------------+
```

Behavior details:

- The name field accepts a file name only, not a full path.
- The editor combo is based on settings, the computer name, and the file extension currently typed in the Name field.
- Invalid names show inline validation and keep the dialog open.
- Existing files require explicit overwrite confirmation.
- On success, the new file becomes the focused item.

### `cmd/pane/contextMenuCurrentDirectory`

User-facing feature: open the Windows shell context menu for the folder currently shown in the active pane.

Flow:

1. User navigates to `C:\Work`.
2. User invokes `Context Menu for Current Directory`.
3. RedSalamander opens the shell context menu for `C:\Work`.
4. If a shell action may have changed contents, the pane refreshes after the menu closes.

Low-fi menu target:

```text
Active pane: C:\Work

Shell context menu target:
  C:\Work

Not target:
  focused file inside C:\Work
```

Behavior details:

- This command is different from opening the context menu for selected items.
- It targets the directory container itself.
- Unsupported virtual/plugin directories show a localized unsupported message.

### `cmd/pane/openSecurity`

User-facing feature: open the Windows Security tab for the focused item.

Flow:

1. User focuses a file or directory.
2. User invokes `Open Security`.
3. Windows opens the Properties dialog on the Security tab.

Low-fi external UI:

```text
RedSalamander
  focused item: C:\Work\report.docx

Windows Properties
  tab: Security
  object: report.docx
```

Behavior details:

- If no item is focused, the active directory is used.
- RedSalamander does not reimplement the Windows ACL editor.
- Plugin file systems are unsupported unless they expose a local shell path.

### `cmd/pane/changeAttributes`

User-facing feature: inspect/update Windows file attributes and optionally remove alternate data streams from selected local items.

Flow:

1. User selects one or more files/directories.
2. User invokes `Change Attributes`.
3. RedSalamander shows current attribute state and whether extra streams were found or can be scanned.
4. User changes attributes and optionally checks `Remove alternate data streams`.
5. RedSalamander applies the requested work.
6. RedSalamander shows a report of changed attributes, removed streams, skipped items, and failures.
7. RedSalamander refreshes affected rows.

Low-fi dialog:

```text
+--------------------------------------+
| Change attributes                    |
|                                      |
| Items: 3 selected                    |
|                                      |
| [~] Read-only                        |
| [ ] Hidden                           |
| [x] Archive                          |
| [ ] System                           |
|                                      |
| [ ] Remove alternate data streams    |
|     Scan selected files before apply |
|                                      |
| [Apply] [Cancel]                     |
+--------------------------------------+
```

Low-fi report:

```text
+-----------------------------------------------+
| Change attributes report                       |
|                                               |
| Attributes changed: 3 items                    |
| Streams removed:     2 streams from 1 item     |
| Skipped:             0                         |
| Failed:              1                         |
|                                               |
| [Copy Report] [Close]                          |
+-----------------------------------------------+
```

Behavior details:

- Mixed multi-selection state is shown as `[~]`.
- Stream removal is opt-in and local-file-system only.
- The report must include stream names where Windows exposes them.
- Failed items are summarized after apply.
- The dialog does not hide access-denied failures.

### `cmd/pane/goToShortcutOrLinkTarget`

User-facing feature: jump from a shortcut, link, reparse point, junction, or mount point to its target, and expose that target in file properties.

Flow:

1. User focuses `Build Logs.lnk`, a `.url` file, a junction, a symlink, or a mount point.
2. User invokes `Go To Shortcut or Link Target`.
3. RedSalamander resolves the target.
4. If the target is a local file, RedSalamander opens the containing folder and focuses the file.
5. If the target is a local folder, RedSalamander navigates to that folder.

Low-fi navigation:

```text
Before:
  C:\Desktop
  > Build Logs.lnk

Resolved target:
  D:\Logs\build-2026.txt

After:
  D:\Logs
  > build-2026.txt
```

Behavior details:

- Broken shortcuts show a localized message.
- `.url` files follow the URL policy documented in the UI spec.
- Reparse points, junctions, symlinks, and mount points are resolved without following them blindly during enumeration.
- File Properties gains a link/target section for `.lnk`, `.url`, reparse point, junction, and mount point entries.
- The command does not modify the shortcut.

### `cmd/pane/newFromShellTemplate`

User-facing feature: create a new file using Windows ShellNew templates.

Flow:

1. User opens `Files > New`.
2. RedSalamander shows ShellNew template types.
3. User chooses a template.
4. RedSalamander prompts for a file name.
5. The file is created, the pane refreshes, and the new item is focused.

Low-fi menu and dialog:

```text
Files
  New >
      Text Document
      Bitmap Image
      Rich Text Document

+--------------------------------------+
| New Text Document                    |
|                                      |
| Folder: C:\Work                      |
| Name:   [ New Text Document.txt    ] |
|                                      |
| [Create] [Cancel]                    |
+--------------------------------------+
```

Behavior details:

- Template list comes from the ShellNew registry view.
- Unsafe template commands are not invoked.
- Empty template list disables the menu and shows a clear shortcut message.

### `cmd/pane/newFromShellTemplate/<templateId>`

User-facing feature: create a file from one exact ShellNew template.

Flow:

1. Shortcut dispatch resolves the template id.
2. RedSalamander validates that the template still exists.
3. RedSalamander opens the same file-name prompt used by the menu path.

No custom UI is shown before the file-name prompt unless the template id is stale or invalid.

### `cmd/pane/clipboardCut`

User-facing feature: cut selected files for a later move, while preserving text-edit cut behavior.

Flow:

1. If a navigation edit field owns focus, `Cut` cuts selected text.
2. Otherwise, selected local files are placed on the clipboard with move intent.
3. Status text can briefly show how many items were cut.
4. A later paste operation moves those files according to existing paste behavior.

Low-fi status:

```text
Active pane: C:\Work
Selected: 3 files

Status: 3 items marked for move
```

Behavior details:

- No file is moved until paste happens.
- Unsupported plugin selections are not placed on the shell clipboard.
- The command must not silently fall through when focus is outside an edit control.

### `cmd/pane/clipboardPasteShortcut`

User-facing feature: create shortcut files in the active directory for clipboard file paths.

Flow:

1. User copies one or more files in RedSalamander or Explorer.
2. User navigates to the destination directory.
3. User invokes `Paste Shortcut`.
4. RedSalamander creates `.lnk` files pointing to the clipboard paths.
5. Created shortcuts are focused or selected.

Low-fi result:

```text
Clipboard:
  C:\Tools\build.exe

Destination pane before:
  D:\Launchers

Destination pane after:
  D:\Launchers
  > build - Shortcut.lnk
```

Behavior details:

- Name collisions use the existing unique-name style.
- Invalid clipboard contents produce "The clipboard does not contain files that can be used for shortcuts."
- Navigation edit controls keep text paste behavior.

### `cmd/pane/quickSearch`

User-facing feature: use the pane's integrated incremental search to highlight matches and navigate among folder items.

Flow:

1. User invokes `Quick Search`.
2. The active pane enters integrated search mode.
3. User types.
4. All matching names are highlighted.
5. The best item whose name starts with the search text is selected.
6. Keyboard navigation moves through matching items without leaving the folder list.
7. Enter accepts the current item; Escape clears search mode.

Low-fi pane:

```text
+------------------------------------------------+
| C:\Work                                        |
| Name                 Size       Modified       |
| report.docx          32 KB      2026-05-01  *  |
| readme.md            4 KB       2026-04-28  >  |
| results.csv          91 KB      2026-04-27  *  |
|                                                |
| Search: re                                     |
+------------------------------------------------+
```

Behavior details:

- Quick search is temporary, pane-scoped, and integrated with folder item navigation.
- It does not replace the persistent filter bar.
- Highlighting applies to every match, while selection moves to the preferred starts-with match.
- Up/Down or the configured next/previous-match keys navigate between matched items.
- No-match state is visible but non-modal.

### `cmd/pane/bringCurrentDirToCommandLine`

User-facing feature: insert the active directory into the command-line input.

Flow:

1. User invokes `Bring Current Directory to Command Line`.
2. RedSalamander shows/focuses command-line input.
3. The active directory path is inserted with proper quoting.

Low-fi input:

```text
+------------------------------------------------+
| Command: [ "C:\src\RedSalamander"          ]   |
+------------------------------------------------+
```

Behavior details:

- The command-line input is distinct from quick search mode.
- The input appears above the function bar, receives keyboard focus, and is associated with the invoking pane.
- Insertions append at the caret or replace the edit selection, adding a single separating space when adjacent text would otherwise touch.
- `Enter` launches the command through the system command processor with the pane's current local folder as working directory, then clears and hides the input on success.
- `Escape` hides the input and restores folder-view focus.

### `cmd/pane/bringFilenameToCommandLine`

User-facing feature: insert the focused file name or selected item path into command-line input.

Flow:

1. User focuses `build.ps1`.
2. User invokes `Bring Filename to Command Line`.
3. Command-line input opens and receives the file name/path with proper quoting.

Low-fi input:

```text
Focused item:
  build.ps1

Command line:
  [ "build.ps1"                                  ]
```

Behavior details:

- With no explicit selection, the focused display name is inserted.
- With an explicit selection, full local item paths are inserted individually. If the focused item is selected, it is inserted first and the remaining selected items keep pane order.
- Empty panes show a localized no-item message.

### `cmd/pane/userMenu`

User-facing feature: run user-configured external commands from a dynamic menu backed by a settings page.

Flow:

1. User opens `Commands > User Menu`.
2. RedSalamander shows configured commands.
3. User chooses one.
4. Macros are expanded using the active pane context.
5. RedSalamander launches the command.

Low-fi menu:

```text
Commands
  User Menu >
      Build Current Project
      Open Here in Terminal
      Compare With Tool
      ----------------
      Configure User Menu...
```

Empty state:

```text
+------------------------------------------+
| User Menu is empty                       |
|                                          |
| Add commands in Preferences to use this. |
|                                          |
| [Open Preferences] [Close]               |
+------------------------------------------+
```

Behavior details:

- User Menu has its own settings page and schema entries.
- User Menu external commands use the same macro engine as View, Alternate View, Edit, Alternate Edit, View With, and Edit With.
- Macros are visible in preferences so users can understand what will be launched.
- Missing executables disable individual entries.
- Computer-specific entries can be hidden or disabled when they do not apply to the current computer.
- Launch failures identify the entry name.

### `cmd/pane/userMenu/<itemId>`

User-facing feature: run one exact user menu item from a shortcut.

Flow:

1. Shortcut dispatch resolves the user menu item id.
2. RedSalamander validates the entry.
3. Macro expansion and launch use the same path as the menu command.

No custom UI is shown unless validation or launch fails.

### `cmd/pane/viewOptions/toggleFileExtensions`

User-facing feature: show or hide file extensions in the target pane.

Flow:

1. User invokes the toggle.
2. The pane redraws names immediately.
3. Menu check state updates.
4. Setting persists.

Low-fi before/after:

```text
Extensions shown:
  report.docx
  notes.txt
  archive.tar.gz

Extensions hidden:
  report
  notes
  archive.tar
```

Behavior details:

- This is display-only.
- Rename, copy, delete, sorting keys, and actual paths continue using real names.
- Folder names are unchanged.

### `cmd/pane/viewOptions/toggleThumbnails`

User-facing feature: switch the pane from normal file icons to thumbnail-capable visuals.

Flow:

1. User toggles `Thumbnails`.
2. Visible rows keep icon fallback immediately.
3. Thumbnail requests start asynchronously for visible items.
4. Rows update as thumbnails arrive.
5. multiple size of thumbnail are available

Low-fi pane:

```text
Thumbnail mode

+----------------------+           +----------------------+           
|                      |           |                      |
|                      |           |                      |
| image                |           | image                |
|                      |           |                      |
|                      |           |                      |
+----------------------+           +----------------------+
vacation.jpg      3.1 MB           vacation2.jpg      3.1 MB 
+----------------------+           +----------------------+           
|                      |           |                      |
|                      |           |                      |
| image                |           | image                |
|                      |           |                      |
|                      |           |                      |
+----------------------+           +----------------------+
 design.png      412 KB            design2.png      412 KB 
+----------------------+  
|                      |  
|                      |  
| [icon]               |  
|                      |  
|                      |  
+----------------------+
  readme.txt          4 KB
  ```

Behavior details:

- The UI stays responsive while thumbnails load.
- Stale thumbnail results from previous folders are ignored.
- Large folders must have archived perf evidence.

### `cmd/pane/viewOptions/togglePreviewPane`

User-facing feature: use the opposite pane as a tabbed preview host for the focused item in the main pane.

Flow:

1. User toggles `Preview Pane`.
2. The opposite pane gains top tabs: `Folder` and `Preview`.
3. The `Preview` tab displays the focused item from the main pane.
4. Changing selection in the main pane updates the preview.
5. User can switch the opposite pane back to its `Folder` tab without closing preview mode.
6. Closing preview removes the `Preview` tab and restores the opposite pane to normal folder mode.

Low-fi layout:

```text
Main menu
+--------------------------------------------------------------+
| Left pane: Folder view          | Right pane tabs             |
| C:\Work                         | [Folder] [Preview]          |
|                                 | Nav: C:\Work                |
| > notes.txt                     | --------------------------  |
|   report.docx                   | Meeting notes...            |
|   image.png                     |                              |
|                                 | Status: notes.txt, 4 KB      |
+--------------------------------------------------------------+
Function bar, when visible
```

Behavior details:

- The preview lives in the other pane, not as a small side panel inside the same pane.
- The tab strip appears just under the main menu area for that pane.
- Navigation bar and status bar are part of the preview host.
- The preview area extends down to the function bar, or to the bottom of the window when the function bar is hidden.
- Preview loading is bounded or asynchronous.
- Navigation cancels stale preview loads.
- The preview pane does not steal folder-view keyboard navigation.

### `cmd/pane/viewOptions/toggleFilterBar`

User-facing feature: show a persistent filter bar for the target pane.

Flow:

1. User toggles `Filter Bar`.
2. A filter input appears in the pane chrome.
3. Existing filter state appears in the input.
4. Typing updates the pane using existing filter semantics.

Low-fi pane:

```text
+------------------------------------------------+
| C:\Work                                        |
| Filter: [ *.cpp                            ] x |
|                                                |
| main.cpp                                       |
| CommandRegistry.cpp                            |
+------------------------------------------------+
```

Behavior details:

- Closing the bar does not silently clear the filter unless the UI spec explicitly chooses that behavior.
- The command is different from temporary quick search.

### `cmd/pane/viewOptions/toggleNavigationBar`

User-facing feature: show or hide the address/navigation bar for one pane.

Flow:

1. User toggles `Navigation Bar`.
2. The target pane layout updates immediately.
3. Setting persists.

Low-fi before/after:

```text
Navigation bar visible:
+----------------------------------+
| C:\Work\Project              v   |
| file1.cpp                        |
| file2.cpp                        |
+----------------------------------+

Navigation bar hidden:
+----------------------------------+
| file1.cpp                        |
| file2.cpp                        |
+----------------------------------+
```

Behavior details:

- Address-bar focus commands either show the bar first or explain that it is hidden.
- Left/right menu items target the named pane; shortcut targets active pane.

### `cmd/pane/viewOptions/toggleStatusBar`

User-facing feature: make the generic status-bar shortcut behave like the existing left/right menu commands.

Flow:

1. User focuses a pane.
2. User invokes the generic status-bar command.
3. The active pane status bar toggles.

Low-fi behavior:

```text
Active pane: right
Command: Toggle Status Bar
Result: right status bar hidden/shown
```

Behavior details:

- This is mostly a routing fix.
- Existing explicit left/right menu behavior remains unchanged.

### `cmd/pane/makeFileList`

User-facing feature: create a JSON, text, or CSV listing of files using user-selected fields and optional text macros.

Flow:

1. User invokes `Make File List`.
2. RedSalamander opens an options dialog.
3. User chooses scope, fields, format, text macro template, and output.
4. RedSalamander generates the list.
5. Result is copied to clipboard or written to a file.
6. The selected options are saved in settings for the next use.

Low-fi dialog:

```text
+----------------------------------------------+
| Make file list                               |
|                                              |
| Source:  (o) Selected items  ( ) Current dir |
| [ ] Recursive                                |
|                                              |
| Fields:  [x] Name [x] Size [x] Modified      |
|          [ ] Attributes [ ] Full path        |
|                                              |
| Format:  (o) Text  ( ) CSV  ( ) JSON         |
| Text macro:                                  |
| [ {fullPath}\t{size}\t{modified}          ] |
|                                              |
| Output:  (o) Clipboard  ( ) Save to file     |
|                                              |
| [Create] [Cancel]                            |
+----------------------------------------------+
```

Behavior details:

- Last picked source, fields, format, macro, and output mode are persisted in settings.
- JSON output uses a documented schema.
- Large recursive output runs with progress and cancellation.
- CSV output quotes commas, quotes, and newlines correctly.
- Text output expands the same path/name/date/size field macros documented for this feature.
- Plugin panes require enumeration fields; unsupported providers show a clear message.

### `cmd/pane/listOpenedFiles`

User-facing feature: inspect files currently opened through RedSalamander.

Flow:

1. User invokes `List Opened Files`.
2. RedSalamander opens a list dialog.
3. User can select an entry and focus its owning pane item when possible.

Low-fi dialog:

```text
+----------------------------------------------------------+
| Opened files                                             |
|                                                          |
| File                 Source        Opened by             |
| notes.txt            C:\Work       Text Viewer           |
| report.docx          C:\Docs       External editor       |
| image.png            C:\Images     Preview Pane          |
|                                                          |
| [Focus Item] [Close]                                     |
+----------------------------------------------------------+
```

Behavior details:

- Empty state is explicit: "No files are currently opened by RedSalamander."
- Closed viewers/editors are removed from the list.

### `cmd/pane/shares`

User-facing feature: show local Windows shared directories.

Flow:

1. User invokes `Shared Directories`.
2. RedSalamander enumerates local shares.
3. A dialog displays share name, path, type, and remark.
4. User can open Windows share management if available.

Low-fi dialog:

```text
+----------------------------------------------------------+
| Shared directories                                       |
|                                                          |
| Share name    Path                  Type      Remark     |
| Public        C:\Users\Public       Disk      Public     |
| Builds        D:\Builds             Disk      CI output  |
|                                                          |
| [Open Management] [Close]                                |
+----------------------------------------------------------+
```

Behavior details:

- Access denied is non-fatal and visible.
- This command lists shares; it does not create or edit shares in RedSalamander.

### `cmd/app/rereadAssociations`

User-facing feature: refresh file associations, plugin-driven menus, viewer/editor mappings, and icon/thumbnail caches.

Flow:

1. User invokes `Reread Associations`.
2. RedSalamander reloads associations.
3. Dynamic menus and visible panes refresh.
4. User sees a short success or partial-failure notification.

Low-fi notification:

```text
+--------------------------------------------+
| Associations refreshed                      |
| Viewers, editors, plugins, and icons reread |
+--------------------------------------------+
```

Behavior details:

- The command preserves valid active selections.
- If one provider fails, RedSalamander keeps previous valid state where possible and reports partial failure.

### `cmd/pane/pack`

User-facing feature: create an archive from selected files/directories.

Flow:

1. User selects files or focuses a directory.
2. User invokes `Pack`.
3. RedSalamander opens archive options.
4. User confirms.
5. File-operation progress starts.
6. The archive appears in the destination and the pane refreshes.

Low-fi dialog:

```text
+------------------------------------------------+
| Pack                                           |
|                                                |
| Source: 3 selected items                       |
| Archive: [ C:\Work\selected.zip             ]  |
| Format:  [ ZIP v ]                             |
| Level:   [ Normal v ]                          |
|                                                |
| [Pack] [Cancel]                                |
+------------------------------------------------+
```

Low-fi progress:

```text
Packing selected.zip
[##########--------------] 42%
Current: assets\logo.png
[Cancel]
```

Behavior details:

- `.zip` is the minimum shippable format.
- `.7z` is added only if write/update support is implemented safely.
- Cancellation and overwrite behavior must match the File Operations spec.

### `cmd/pane/unpack`

User-facing feature: extract one or more archives.

Flow:

1. User selects an archive.
2. User invokes `Unpack`.
3. RedSalamander asks for destination and overwrite policy.
4. User confirms.
5. File-operation progress starts.
6. Destination pane refreshes after extraction.

Low-fi dialog:

```text
+------------------------------------------------+
| Unpack                                         |
|                                                |
| Archive: C:\Downloads\source.zip               |
| Destination: [ C:\Downloads\source          ]  |
| Overwrite:   [ Ask v ]                         |
|                                                |
| [Unpack] [Cancel]                              |
+------------------------------------------------+
```

Behavior details:

- Unsupported archive formats show a localized message.
- Partial output from cancellation follows the File Operations cleanup policy.
- Multiple selected archives use a documented destination folder rule.

### `cmd/app/theme/select/<themeId>`

User-facing feature: select a theme from shortcut, command palette, or dynamic menu using an exact theme id.

Flow:

1. Command parameter resolves to a built-in or custom theme.
2. RedSalamander applies the theme.
3. Theme menu check state and visible UI update.

No custom UI is shown unless the theme id is missing.

### `cmd/app/theme/selectNext`

User-facing feature: cycle to the next available theme without opening Preferences.

Flow:

1. User invokes `Next Theme` from the main menu or function-bar shortcut.
2. RedSalamander finds the current theme in the same order shown by the Theme menu.
3. RedSalamander applies the next theme, wrapping from the last theme to the first.
4. Theme menu check state and visible UI update.

Low-fi menu:

```text
View
  Theme >
      Previous Theme
      Next Theme
      ----------------
      Light
      Dark       (checked)
      High Contrast
      Custom Blue
```

Behavior details:

- Built-in and custom themes participate in the same cycle.
- Disabled or invalid custom themes are skipped.
- The function-bar shortcut is added beside the existing theme/view shortcuts.

### `cmd/app/theme/selectPrev`

User-facing feature: cycle to the previous available theme without opening Preferences.

Flow:

1. User invokes `Previous Theme` from the main menu or function-bar shortcut.
2. RedSalamander finds the current theme in the same order shown by the Theme menu.
3. RedSalamander applies the previous theme, wrapping from the first theme to the last.
4. Theme menu check state and visible UI update.

Behavior details:

- Uses the same theme order and skip rules as `cmd/app/theme/selectNext`.
- Function-bar shortcut coverage is required.

### `cmd/app/plugins/toggleEnabled/<pluginId>`

User-facing feature: enable or disable one exact plugin from a dynamic command.

Flow:

1. Command parameter resolves to a plugin.
2. RedSalamander toggles enabled state.
3. Plugin menus refresh.
4. If disabling affects active pane state, the user gets a localized explanation or migration action.

No custom UI is shown for a successful toggle.

### `cmd/app/plugins/configure/<pluginId>`

User-facing feature: open configuration UI for one exact plugin.

Flow:

1. Command parameter resolves to a plugin.
2. If the plugin exposes configuration UI, RedSalamander opens it.
3. If not, RedSalamander shows "This plugin has no configuration."

Low-fi message:

```text
+-----------------------------------+
| Plugin has no configuration       |
|                                   |
| FileSystem7z does not provide     |
| configurable settings.            |
|                                   |
| [Close]                           |
+-----------------------------------+
```

### `cmd/pane/selectFileSystemPlugin/<pluginId>`

User-facing feature: switch the active pane to a specific file-system plugin.

Flow:

1. Command parameter resolves to an enabled file-system plugin.
2. RedSalamander changes the active pane provider.
3. Pane content reloads using that provider.
4. Menu check state updates.

No custom UI is shown unless the plugin is unavailable or disabled.

### `cmd/pane/navigatePath/<pathId>`

User-facing feature: navigate the active pane to an exact dynamic path target.

Flow:

1. Command parameter resolves to a hot path, history path, or dynamic path target.
2. RedSalamander navigates the active pane.
3. If the path no longer exists, RedSalamander shows the standard navigation failure message.

No custom UI is shown for successful navigation.

### `cmd/app/openFileExplorerKnownFolder/<knownFolderId>`

User-facing feature: open a known Windows folder in File Explorer from a parameterized shortcut.

Flow:

1. Command parameter resolves to a known folder id.
2. RedSalamander asks Windows for that folder path.
3. File Explorer opens that location.

No custom UI is shown unless Windows cannot resolve or open the folder.

## Architecture

### Command execution model

Add a single command execution layer used by both menu commands and shortcut commands.

Create a small command instance model:

```cpp
struct CommandInstance
{
    std::wstring canonicalId;
    std::optional<std::wstring> parameter;
    std::optional<Pane> targetPane;
};
```

The final shape can differ if the existing code suggests a better name, but the responsibilities must stay the same:

- Resolve canonical string ids, WM command ids, dynamic menu ids, and pane-local menu ids into one command instance.
- Preserve left/right pane intent for commands such as status bar, navigation bar, file extensions, thumbnails, preview pane, and filter bar.
- Preserve parameter payloads for themes, plugins, viewers, editors, shell templates, known folders, and hot paths.
- Route commands through explicit handlers before falling back to `ShowCommandNotImplementedMessage`.
- Keep unsupported context as a visible state: disabled menu item where possible, localized explanation where execution is attempted through a shortcut.

Recommended files:

- `RedSalamander/CommandRegistry.h`
- `RedSalamander/CommandRegistry.cpp`
- `RedSalamander/RedSalamander.cpp`
- Optional focused files:
  - `RedSalamander/CommandExecution.h`
  - `RedSalamander/CommandExecution.cpp`

### FolderWindow command surface

Keep pane behavior on `FolderWindow`, not in the main window switch. Add focused command files rather than growing `RedSalamander.cpp`.

Recommended files:

- `RedSalamander/FolderWindow.h`
- `RedSalamander/FolderWindow.CommandContext.cpp`
- `RedSalamander/FolderWindow.Viewers.cpp`
- `RedSalamander/FolderWindow.Editors.cpp`
- `RedSalamander/FolderWindow.ExternalActions.cpp`
- `RedSalamander/FolderWindow.ShellCommands.cpp`
- `RedSalamander/FolderWindow.CommandLine.cpp`
- `RedSalamander/FolderWindow.ViewOptions.cpp`
- `RedSalamander/FolderWindow.ListsAndShares.cpp`
- `RedSalamander/FolderWindow.ArchiveCommands.cpp`
- `RedSalamander/FolderWindow.UserMenu.cpp`

Common helpers should provide:

- Current pane and opposite pane lookup.
- Focused item path.
- Selected item paths.
- Current directory path.
- Local-file-system capability checks.
- Plugin file-system capability checks.
- UI-thread assertions for window and Direct2D/DxUi operations.

### Settings and persistence

Extend settings in `Common/SettingsStore.h` and `Common/SettingsStore.cpp`.

Add schema-backed settings for:

- External actions shared by View, Alternate View, Edit, Alternate Edit, View With, Edit With, and User Menu:
  - action id
  - action kind: internal viewer, external viewer, external editor, user command
  - display name
  - enabled state
  - computer-name rule
  - extension rule
  - executable path
  - arguments template
  - working directory template
  - show console flag, if implemented
  - run elevated flag, if implemented
  - macro requirements
- Viewer/editor associations:
  - primary view action by extension and optional computer name
  - alternate view action by extension and optional computer name
  - primary edit action by extension and optional computer name
  - alternate edit action by extension and optional computer name
  - fallback primary view action
  - fallback alternate view action
  - fallback primary edit action
  - fallback alternate edit action
- Editors:
  - editor action list presented by `Edit With`
  - default editor action for `Edit New` when extension-specific settings do not match
- User menu:
  - item id
  - display name
  - external action reference or inline external command action
  - menu ordering
  - computer-name rule
- Pane view options:
  - show file extensions
  - show thumbnails
  - preview pane visible
  - filter bar visible
  - navigation bar visible
  - status bar visible, if the existing status setting is not already schema-backed per pane
- Command line:
  - recent command entries
  - recent quick-search entries, if history is part of the final UI
- Extensions:
  - existing `openWithViewerByExtension`
  - migration path into the new primary/alternate view action associations
  - new primary/alternate editor action associations
- Make file list:
  - last source mode
  - last recursive flag
  - last field selection
  - last output format: JSON, text, or CSV
  - last text macro
  - last output target

Follow the yyjson ownership rules from repo guidance:

- Use copying APIs for dynamic string values.
- Use `yyjson_mut_strncpy` plus `yyjson_mut_obj_add` for dynamic keys.
- Do not pass stack or temporary strings to non-copy mutable APIs.

Update `Specs/Core/Core_SettingsStore.md` with the new schema version and defaults.

### UI and localization

All new user-visible strings must be resource-backed.

Update:

- `RedSalamander/resource.h`
- `RedSalamander/RedSalamander.rc`
- `RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.rc`
- Any preferences dialog resource files touched by the new editor or user-menu UI

Menu labels must remove `[todo]` only after the command is implemented and tested.

## Phase 0: Baseline Reconciliation

### Tasks

- Add a command inventory selftest that classifies each command as:
  - implemented
  - context-disabled
  - intentionally not implemented with fallback
  - stale label/spec marker
- Make the selftest fail if a command has a handler but still has a `[todo]` menu label.
- Make the selftest fail if a registry command has a WM id and reaches the fallback from a normal enabled menu path.
- Retire `cmd/pane/editWidth` from exposed command surfaces:
  - remove or disable menu label
  - remove default shortcut binding
  - remove command registry entry, or keep only a documented compatibility mapping for imported old shortcut settings
  - update specs so it is not listed as a missing command
- Add explicit tests for already implemented stale-marker commands:
  - `cmd/app/viewWidth` remains the splitter view-width adjust command and remains documented that way
  - permanent delete command registration
  - select same name and unselect same name registration
  - sort by attributes registration
  - plugin manager command routing
  - theme select command routing
  - window menu command routing
  - menu bar and function bar toggle commands
  - right-pane hot paths command routing
- Update only markers proven stale by tests.

### Files

- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Shortcuts.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Navigation.cpp`
- Optional new family: `RedSalamander/SelfTest/Commands/Commands.SelfTest.PlannedCommands.cpp`
- `RedSalamander/CommandRegistry.cpp`
- `RedSalamander/ShortcutDefaults.cpp`
- `RedSalamander/RedSalamander.vcxproj`
- `RedSalamander/RedSalamander.vcxproj.filters`
- `Specs/UI/UI_CommandMenuKeyboard.md`
- `RedSalamander/RedSalamander.rc`
- `RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.rc`

### Validation

- `.\build.ps1 -ProjectName RedSalamander`
- `.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "command_registry|shortcut|view_command|planned_command" -TimeoutMultiplier 2`

## Phase 1: Command Routing Foundation

### Tasks

- Add canonical parsing for parameterized command ids.
- Add command registry, resource strings, menu entries, and function-bar shortcut defaults for:
  - `cmd/app/theme/selectNext`
  - `cmd/app/theme/selectPrev`
- Add route support for dynamic menu ids:
  - themes
  - plugins
  - file-system plugin selection
  - hot paths
  - path history
  - known File Explorer folders
  - viewers
  - editors
  - shell templates
  - user menu items
- Make `DispatchShortcutCommand(...)` call the same execution layer as `OnMainWindowCommand(...)`.
- Make pane-scoped shortcut commands target the active pane by default.
- Add helper methods to map generic pane commands to left/right WM ids where the menu model already uses pane-specific ids.
- Fix generic routing for `cmd/pane/viewOptions/toggleStatusBar`.
- Preserve existing behavior for:
  - navigation path menu targets
  - custom theme menu ids
  - plugin menu ids
  - hot path menu ids
  - built-in theme ids
- Ensure unknown commands still show `ShowCommandNotImplementedMessage`.

### Files

- `RedSalamander/CommandRegistry.h`
- `RedSalamander/CommandRegistry.cpp`
- `RedSalamander/RedSalamander.cpp`
- `RedSalamander/ShortcutDefaults.cpp`
- `RedSalamander/resource.h`
- `RedSalamander/RedSalamander.rc`
- `RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.rc`
- Optional: `RedSalamander/CommandExecution.h`
- Optional: `RedSalamander/CommandExecution.cpp`

### Tests

- Parameterized shortcut id canonicalization:
  - `cmd/pane/hotPath/<id>`
  - `cmd/pane/setHotPath/<id>`
  - `cmd/pane/goDriveRoot/<drive>`
  - `cmd/app/theme/select/<themeId>`
  - `cmd/app/theme/selectNext`
  - `cmd/app/theme/selectPrev`
  - `cmd/pane/selectFileSystemPlugin/<pluginId>`
  - `cmd/app/openFileExplorerKnownFolder/<knownFolderId>`
- Generic pane command routing:
  - status bar uses active pane
  - left/right menu ids continue targeting their explicit pane
- Unknown command still produces the fallback and does not crash.

### Validation

- `.\build.ps1 -ProjectName RedSalamander`
- `.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_parameterized|cmd_shortcut_dispatch|cmd_view_options_statusbar" -TimeoutMultiplier 2`

## Phase 2: Viewer and Editor Commands

### Behavior

#### View With

`cmd/pane/viewWith` displays the list of available viewer actions for the focused file and lets the user pick one.

Required behavior:

- The Files menu contains a dynamic `View With` submenu populated from internal viewer plugins and external viewer actions in settings.
- The command is enabled when the active pane has a focused regular file.
- The command is disabled for directories, virtual entries that cannot provide a stream, and unsupported plugin file-system entries.
- `cmd/pane/viewWith/<viewerId>` works from shortcuts.
- The default `View` command keeps its configured primary view action.
- Viewer action failure shows a localized error with the action display name and file name.
- Filtering honors extension rules, computer-name rules, enabled state, and executable availability.

Implementation notes:

- Reuse `OpenViewerWithPlugin(...)` and `ViewerOpenContext`.
- Add a shared external-action launch helper for external viewer actions.
- Do not duplicate viewer instance ownership rules.
- Keep action ids stable and use display names only for UI.

#### Alternate View

`cmd/pane/alternateView` opens the focused file using the configured alternate view action.

Required behavior:

- Alternate View is configured in settings near the primary View action.
- The alternate action can be an internal viewer plugin or an external program.
- Resolution honors extension rules, computer-name rules, and fallback action rules.
- External program actions support macros for path, full path, path plus file name, file name, selected paths file, and computer name.
- If no alternate view action is configured, show a localized "no alternate viewer configured" message rather than the generic not-implemented alert.

#### Edit

`cmd/pane/edit` opens the focused file using the configured primary edit action.

Required behavior:

- The primary edit action is configured by extension, optional computer name, and fallback settings.
- The editor action can be an external program with macros:
  - `{path}`
  - `{fullPath}`
  - `{pathAndFilename}`
  - `{filename}`
  - `{selectedPathsFile}`
  - `{computer}`
- Launch uses `ShellExecuteExW` with WIL-managed handles.
- Failure shows a localized error with the editor display name and Win32 error.
- Folder refresh is requested after the editor process starts.

#### Edit With

`cmd/pane/editWith` displays the list of available editor actions for the focused file and lets the user pick one.

Required behavior:

- The Files menu contains a dynamic `Edit With` submenu populated from editor actions in settings.
- `cmd/pane/editWith/<editorId>` works from shortcuts.
- The submenu has an entry to open Editors preferences.
- Missing editor executable disables that editor item and shows a reason in preferences.
- Filtering honors extension rules, computer-name rules, enabled state, and executable availability.

#### Alternate Edit

`cmd/pane/alternateEdit` opens the focused file using the configured alternate edit action.

Required behavior:

- Alternate Edit is configured in settings near the primary Edit action.
- The alternate action can be an external program with the same macro support as Edit.
- Resolution honors extension rules, computer-name rules, and fallback action rules.
- If no alternate edit action is configured, show a localized "no alternate editor configured" message.
- Add the command to registry, menus, default shortcut lists if desired, resources, and selftests.

#### Edit New

`cmd/pane/editNew` creates a new file in the active directory and opens it in the editor selected by the user.

Required behavior:

- Prompt for the new file name.
- Infer the extension as the user types.
- Display an `Editor` combo with editor actions that match the inferred extension and current computer.
- Default the combo to the configured primary editor action for that extension/computer.
- Reject empty names, absolute paths, parent traversal, invalid path characters, and collisions unless the user explicitly confirms overwrite.
- Create the file with WIL/Win32 RAII and non-throwing filesystem checks.
- Refresh the active pane and focus the new file.
- Open the new file in the selected editor action.

### Files

- `RedSalamander/FolderWindow.h`
- `RedSalamander/CommandRegistry.cpp`
- `RedSalamander/ShortcutDefaults.cpp`
- `RedSalamander/FolderWindow.Viewers.cpp`
- `RedSalamander/FolderWindow.Editors.cpp`
- `RedSalamander/FolderWindow.ExternalActions.cpp`
- `RedSalamander/Preferences.Editors.h`
- `RedSalamander/Preferences.Editors.cpp`
- `RedSalamander/resource.h`
- `RedSalamander/RedSalamander.rc`
- `RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.rc`
- `Common/SettingsStore.h`
- `Common/Common/SettingsStore.cpp`
- `Specs/Core/Core_SettingsStore.md`
- `Specs/UI/UI_CommandMenuKeyboard.md`
- `Docs/Preferences.md`

### Tests

- Viewer picker population from registered internal viewers and settings-backed external viewer actions.
- Shared external-action macro expansion for `{path}`, `{fullPath}`, `{pathAndFilename}`, `{filename}`, `{selectedPathsFile}`, and `{computer}`.
- `View With` opens the requested internal or external viewer action id.
- `Alternate View` runs the configured alternate view action and does not use automatic next-viewer cycling.
- Missing alternate viewer reports the specific localized message.
- Editor/action settings parse/write roundtrip.
- Primary editor fallback is stable.
- `Edit With` parameterized command launches the selected editor action in dry-run selftest mode.
- `Alternate Edit` registry, menu, shortcut dispatch, settings resolution, and dry-run launch.
- `Edit New` validation rejects unsafe names.
- `Edit New` editor combo filters by typed extension and computer name.
- `Edit New` creates and focuses a file in a temp folder.
- `cmd/pane/editWidth` is retired or compatibility-mapped and is not exposed as a new implemented command.

### Validation

- `.\build.ps1 -ProjectName RedSalamander`
- `.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_view_with|cmd_alternate_view|cmd_editors|settings_editors" -TimeoutMultiplier 2`

## Phase 3: Shell and Item Utility Commands

### Behavior

#### Current Directory Context Menu

`cmd/pane/contextMenuCurrentDirectory` opens the shell context menu for the active pane directory.

Required behavior:

- Use COM shell interfaces with `wil::com_ptr`.
- Initialize COM on the calling thread as required by the existing UI-thread model.
- The menu targets the directory itself, not the focused item.
- Shell verbs that change files trigger pane refresh after menu dismissal.
- Failures show localized shell error text.

#### Open Security

`cmd/pane/openSecurity` opens the Windows Security property page for the focused item or current directory.

Required behavior:

- Use `SHObjectProperties(..., SHOP_FILEPATH, ..., L"Security")` or the equivalent project-approved shell API.
- Support files and directories on the local file system.
- Disable for plugin file systems unless a plugin capability is added.
- Show a localized unsupported-context message for shortcuts.

#### Change Attributes

`cmd/pane/changeAttributes` opens a dialog for file attributes and optional alternate data stream removal.

Required behavior:

- Support single and multi-selection for local files and directories.
- Display mixed state for multi-selection.
- Support read-only, hidden, archive, system where permitted.
- Apply changes with `SetFileAttributesW`.
- Add an opt-in checkbox to remove alternate data streams from selected local files.
- Enumerate streams with Win32 stream APIs and remove only streams that belong to selected files.
- Show a completion report listing changed attributes, removed streams, skipped items, and failures.
- Use non-throwing filesystem checks.
- Log unexpected failures with `Debug::ErrorWithLastError(...)`.
- Refresh affected pane entries after apply.

#### Go To Shortcut or Link Target

`cmd/pane/goToShortcutOrLinkTarget` resolves the target of a shortcut, link, reparse point, junction, or mount point and navigates to it.

Required behavior:

- Support `.lnk` with `IShellLinkW` and `IPersistFile`.
- Support `.url` files by reading the `InternetShortcut` URL.
- Support reparse points, junctions, symbolic links, and mount points with target-resolution helpers that avoid recursive traversal hazards.
- For local file targets, navigate to the containing directory and focus the target item.
- For local directory targets, navigate directly to the directory.
- For URLs, use the default browser or show a confirmation if the project has a URL-opening policy.
- Extend File Properties to show shortcut/link/reparse/mount-point target information.
- Show localized messages for broken links and unsupported targets.

#### New From Shell Template

`cmd/pane/newFromShellTemplate` creates a new file from Windows ShellNew templates.

Required behavior:

- Enumerate ShellNew-capable file types from HKCR/HKCU merged registry view.
- Populate a dynamic submenu.
- Support at least these ShellNew forms:
  - `NullFile`
  - `FileName`
  - `Data`
  - `Command`, only if the shell requires it and the command is safe to invoke through ShellExecute
- Prompt for a file name using the extension's display name.
- Create unique names and focus the created item.
- Keep registry handles RAII-managed.

### Files

- `RedSalamander/FolderWindow.h`
- `RedSalamander/FolderWindow.ShellCommands.cpp`
- `RedSalamander/FolderWindow.ItemProperties.cpp`
- `RedSalamander/RedSalamander.cpp`
- `RedSalamander/resource.h`
- `RedSalamander/RedSalamander.rc`
- `RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.rc`
- `Specs/UI/UI_CommandMenuKeyboard.md`

### Tests

- Current directory context command resolves the active directory path in selftest shell dry-run mode.
- Security command targets focused file, selected directory, and active directory fallback.
- Attribute dialog model handles single-selection and mixed multi-selection.
- Attribute apply changes temp file attributes and restores them.
- Attribute command removes alternate data streams from temp files only when the option is checked.
- Attribute report lists changed items, removed streams, skipped items, and failures.
- `.lnk` target resolution focuses the target.
- Broken `.lnk` reports localized failure.
- `.url` parsing handles local and web targets.
- Reparse point, junction, and mount point target resolution.
- File Properties shows shortcut/link/reparse/mount-point target information.
- ShellNew enumeration handles synthetic registry data through a test seam or in-memory provider.
- ShellNew file creation rejects invalid names and creates the requested extension.

### Validation

- `.\build.ps1 -ProjectName RedSalamander`
- `.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_shell_context|cmd_security|cmd_attributes|cmd_shortcut_target|cmd_shell_new" -TimeoutMultiplier 2`

## Phase 4: Clipboard, Command Line, Quick Search, and User Menu

### Behavior

#### Clipboard Cut

`cmd/pane/clipboardCut` must cut selected files from the active pane when focus is not inside an edit control.

Required behavior:

- Preserve existing edit-control handling in `NavigationView`.
- For folder selections, place `CF_HDROP` plus Preferred DropEffect `DROPEFFECT_MOVE` on the clipboard.
- Disable when the selection cannot be represented as local shell paths.
- Use WIL-managed global memory wrappers or an existing project clipboard helper.
- Keep clipboard ownership exception-safe without `catch (...)`.

#### Paste Shortcut

`cmd/pane/clipboardPasteShortcut` creates shortcuts in the active pane directory for clipboard files.

Required behavior:

- Preserve edit-control paste behavior where applicable.
- Read `CF_HDROP` from the clipboard.
- For each source path, create a `.lnk` in the current directory.
- Generate unique names using the same naming style as copy/paste collision handling.
- Use `IShellLinkW` and `IPersistFile` with `wil::com_ptr`.
- Refresh and focus created shortcuts.

#### Quick Search and Command-Line Input

Quick search must use the pane's integrated search behavior. Command-line input remains a separate command host.

Required behavior:

- `cmd/pane/quickSearch` activates integrated pane search.
- While typing, all matches are highlighted.
- The preferred starts-with match is selected.
- Keyboard navigation moves between matched folder items without leaving the folder view.
- `cmd/pane/bringCurrentDirToCommandLine` inserts the active directory path into command-line input.
- `cmd/pane/bringFilenameToCommandLine` inserts the focused item name or selected item path into command-line input.
- Escape closes or clears the host according to the mode.
- Enter in quick search mode accepts the selected matched item.
- Enter in command line mode executes through `ShellExecuteExW` with current directory as working directory.
- The command line must not block the UI thread while launching.
- Normal folder filtering remains separate from quick search.

#### User Menu

`cmd/pane/userMenu` opens a dynamic menu of configured external commands.

Required behavior:

- User menu entries are stored in settings.
- Add a User Menu settings page.
- Entries support display name, executable, arguments, working directory, enabled state, menu order, and computer-name rule.
- User Menu uses the same macro engine as View, Alternate View, Edit, Alternate Edit, View With, and Edit With.
- Macro expansion supports at least:
  - `{path}`
  - `{fullPath}`
  - `{pathAndFilename}`
  - `{filename}`
  - `{selectedPathsFile}`
  - `{computer}`
  - `{oppositeDir}`
- If a macro needs a temporary selected-paths file, create it in a safe temp location and delete it after launch when possible.
- Empty user menu opens the user menu preferences page with a localized explanatory message.
- Parameterized shortcut `cmd/pane/userMenu/<itemId>` executes the matching item.

### Files

- `RedSalamander/FolderWindow.h`
- `RedSalamander/FolderWindow.CommandLine.cpp`
- `RedSalamander/FolderWindow.ExternalActions.cpp`
- `RedSalamander/FolderWindow.UserMenu.cpp`
- `RedSalamander/Preferences.UserMenu.h`
- `RedSalamander/Preferences.UserMenu.cpp`
- `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`
- `RedSalamander/NavigationView.Edit.cpp`
- `Common/SettingsStore.h`
- `Common/Common/SettingsStore.cpp`
- `Specs/Core/Core_SettingsStore.md`
- `Specs/UI/UI_CommandMenuKeyboard.md`
- `Docs/Preferences.md`

### Tests

- Clipboard cut writes drop files and move effect for local selections.
- Clipboard cut still edits text when navigation edit owns focus.
- Paste shortcut creates unique `.lnk` files for one and multiple paths.
- Paste shortcut rejects unsupported clipboard contents with a localized message.
- Quick search activates integrated pane search, highlights all matches, selects starts-with match, and navigates between matches.
- Bring current directory inserts the active directory path.
- Bring filename inserts the focused item name.
- Command-line execution can be dry-run tested without launching external commands.
- User menu settings roundtrip.
- User menu macro expansion handles paths with spaces and shares implementation with viewer/editor actions.
- User menu settings page persists ordering, command, arguments, working directory, enabled state, and computer-name rule.
- Empty user menu opens preferences path.

### Validation

- `.\build.ps1 -ProjectName RedSalamander`
- `.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_clipboard_cut|cmd_paste_shortcut|cmd_quick_search|cmd_command_line|cmd_user_menu" -TimeoutMultiplier 2`

## Phase 5: Pane View Options

### Behavior

#### File Extensions

`cmd/pane/viewOptions/toggleFileExtensions` toggles display of file extensions for the target pane.

Required behavior:

- The toggle is display-only; operations continue using real paths.
- Folder names are not modified.
- File names without extensions display unchanged.
- Sorting continues to use the configured sort mode, not the stripped display string unless the UI spec explicitly says otherwise.
- The setting persists per pane or globally according to the existing view-options model.

#### Thumbnails

`cmd/pane/viewOptions/toggleThumbnails` toggles thumbnail display in the target pane.

Required behavior:

- Use an asynchronous thumbnail cache.
- Bound visible-work requests to avoid UI stalls while scrolling.
- Fall back to existing icons when a thumbnail is unavailable.
- Cancel or ignore stale thumbnail requests after navigation.
- Integrate with DPI-aware rendering.
- Record performance metrics for thumbnail request count, completed count, cache hits, and visible-frame cost.

#### Preview Pane

`cmd/pane/viewOptions/togglePreviewPane` toggles preview mode in the opposite pane.

Required behavior:

- Preview is hosted in the other pane, not as an inline side panel in the current pane.
- The preview host has top tabs for `Folder` and `Preview`.
- The preview tab appears when preview opens and disappears when preview closes.
- Selection changes in the main pane update the preview in the other pane.
- Navigation bar and status bar belong to the preview host while preview is active.
- The preview area starts below the main menu/tab strip and extends to the function bar, or to the bottom window edge when the function bar is hidden.
- File loading is asynchronous or bounded so selection changes do not block the UI thread.
- Text previews use the existing viewer/text rendering path where practical.
- Unsupported files show a compact localized unsupported preview message inside the Preview tab.
- Preview content is canceled or detached on navigation and pane destruction.
- The layout remains stable on DPI changes, resize, and function-bar visibility changes.

#### Filter Bar

`cmd/pane/viewOptions/toggleFilterBar` toggles a visible filter bar for the target pane.

Required behavior:

- Reuse existing filter command behavior and filter matching.
- The bar shows current filter state.
- Closing the bar does not silently clear the filter unless the UI spec says so.
- Keyboard focus returns to the folder view when expected.

#### Navigation Bar

`cmd/pane/viewOptions/toggleNavigationBar` toggles the navigation/address bar for the target pane.

Required behavior:

- Existing `NavigationView` command handling still works when the bar is visible.
- If hidden, focus shortcuts that require the address bar show it before focusing or show a localized explanation.
- The setting persists.
- Left/right menu items target the selected pane exactly.

#### Generic Status Bar Routing

`cmd/pane/viewOptions/toggleStatusBar` must work from shortcuts against the active pane.

Required behavior:

- Keep existing left/right menu commands.
- Add generic command routing for the active pane.
- Persist and render state exactly as the menu commands do.

### Files

- `RedSalamander/FolderWindow.h`
- `RedSalamander/FolderWindow.ViewOptions.cpp`
- `RedSalamander/FolderWindow.Layout.cpp`
- `RedSalamander/FolderView.h`
- `RedSalamander/FolderView.cpp`
- `RedSalamander/FolderView.Rendering.cpp`
- `RedSalamander/NavigationView.cpp`
- `Common/SettingsStore.h`
- `Common/Common/SettingsStore.cpp`
- `Specs/Core/Core_SettingsStore.md`
- `Specs/UI/UI_CommandMenuKeyboard.md`
- `Specs/Testing/Testing_PerformanceValidation.md`, only if new metric names need to be documented

### Tests

- File-extension toggle changes display labels without changing operation paths.
- Thumbnails toggle schedules bounded async work and renders icon fallback.
- Thumbnail stale-request cancellation is deterministic under selftest.
- Preview pane uses the opposite pane as a tabbed host, follows main-pane selection, manages navigation/status bar layout, and cancels stale loads.
- Filter bar visibility persists and reflects existing filter state.
- Navigation bar visibility persists and does not break address-bar focus command.
- Generic status bar shortcut targets the active pane.

### Perf Evidence

Required protected scenarios:

- Large directory enumeration with thumbnails disabled.
- Large directory enumeration with thumbnails enabled and cold cache.
- Large directory enumeration with thumbnails enabled and warm cache.
- Rapid selection changes with preview pane enabled.
- Resize and DPI-change path with preview pane and filter bar visible.

Archive before/after runs under `Specs/TestRuns/` with:

- scenario definition
- build configuration
- machine hash
- metric summary
- raw selftest/perf output
- notes for regressions or accepted tradeoffs

### Validation

- `.\build.ps1 -ProjectName RedSalamander`
- `.\build.ps1 -ProjectName RedSalamander -Configuration Release`
- `.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_view_options|thumbnail|preview_pane|filter_bar|navigation_bar" -TimeoutMultiplier 2`

## Phase 6: Lists, Associations, Shares, Pack, and Unpack

### Behavior

#### Make File List

`cmd/pane/makeFileList` generates a file list from the active pane in JSON, text, or CSV format.

Required behavior:

- Dialog options:
  - selected items or current directory
  - recursive or non-recursive
  - include files
  - include directories
  - include size
  - include modified time
  - output as JSON, text, or CSV
  - text macro template for custom text output
  - copy to clipboard or save to file
- Persist the last selected options in settings:
  - source mode
  - recursive flag
  - field set
  - output format
  - text macro
  - output target
- JSON output uses a documented schema.
- Text output expands documented file-list macros.
- Use UTF-8 with BOM or UTF-16 according to existing project text-file conventions.
- Large directory work runs off the UI thread with progress and cancellation.
- Plugin file systems are supported only when enumeration APIs provide the needed fields; otherwise show localized unsupported-context text.

#### List Opened Files

`cmd/pane/listOpenedFiles` shows currently opened files.

Required behavior:

- Include open viewer instances tracked by `FolderWindow.Viewers.cpp`.
- Include edit processes launched by this app if editor process tracking is added.
- Include preview-pane file if preview keeps an active file handle.
- The dialog supports focusing the owning pane item when possible.
- Empty state uses localized text.

#### Shared Directories

`cmd/pane/shares` shows shared directories on the local machine.

Required behavior:

- Use `NetShareEnum` with RAII cleanup.
- Display share name, local path, type, and remark where available.
- Provide an action to open Windows shared folders management if appropriate.
- Handle access denied with localized text and a non-fatal warning.

#### Reread Associations

`cmd/app/rereadAssociations` refreshes external associations and dynamic menus.

Required behavior:

- Reload settings-backed extension associations.
- Refresh viewer/editor dynamic menus.
- Refresh file-system plugin selection menus.
- Refresh icon/thumbnail association caches.
- Refresh visible panes after associations are reloaded.
- Preserve existing active plugin selection where still valid.

#### Pack

`cmd/pane/pack` creates an archive from selected items or the active directory.

Required behavior:

- Dialog options:
  - archive path
  - format, at minimum `.zip`; add `.7z` if the existing 7-Zip integration supports write/update APIs
  - compression level
  - include selected items or current directory
  - overwrite policy
- Run through the file-operation progress/cancel infrastructure.
- Support long paths.
- Do not block the UI thread while enumerating or compressing.
- Use deterministic archive entry ordering in tests.
- Show localized errors for unsupported plugin file-system sources.

#### Unpack

`cmd/pane/unpack` extracts selected archives.

Required behavior:

- Support archives readable by the existing FileSystem7z plugin.
- Prompt for destination.
- Run through file-operation progress/cancel infrastructure.
- Preserve directory structure.
- Handle overwrite policy consistently with copy operations.
- Support cancellation and cleanup of partially extracted files according to the File Operations spec.
- Refresh destination pane after completion.

### Files

- `RedSalamander/FolderWindow.h`
- `RedSalamander/FolderWindow.ListsAndShares.cpp`
- `RedSalamander/FolderWindow.ArchiveCommands.cpp`
- `RedSalamander/FolderWindow.FileOperations.cpp`
- `RedSalamander/FolderWindow.Viewers.cpp`
- `Common/SettingsStore.h`
- `Common/Common/SettingsStore.cpp`
- `Plugins/FileSystem7z/*`, if pack/update support is added there
- `Specs/FileSystem/FileSystem_FileOperations.md`
- `Specs/UI/UI_CommandMenuKeyboard.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Docs/Todo.md`

### Tests

- Make file list writes deterministic JSON, text, and CSV output for temp folders.
- Make file list macro expansion and saved option reload are covered.
- Make file list cancellation leaves no partial output unless user selected an explicit output file and overwrite policy permits it.
- List opened files includes viewer instances opened during selftest.
- Shared directories command handles synthetic provider data and access-denied provider result.
- Reread associations reloads editor/viewer extension maps and rebuilds dynamic menus.
- Pack creates a deterministic small archive from temp files.
- Pack cancellation is tested with a controlled large input provider.
- Unpack extracts a small archive and preserves contents.
- Unpack overwrite policy matches copy operation policy.
- Unpack cancellation does not leak progress state or payload messages.

### Perf Evidence

Required protected scenarios:

- Make file list for a large directory.
- Pack many small files.
- Pack one large file.
- Unpack many small files.
- Unpack one large file.
- Reread associations with many extension mappings and plugins.

Archive evidence under `Specs/TestRuns/` with before/after comparisons where an existing comparable path exists.

### Validation

- `.\build.ps1 -ProjectName RedSalamander`
- `.\build.ps1 -ProjectName RedSalamander -Configuration Release`
- `.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_make_file_list|cmd_opened_files|cmd_shares|cmd_reread_associations|cmd_archive" -TimeoutMultiplier 2`
- `.\Tools\Run-AllTests.ps1 -Suite FileOps -CaseFilter "archive|pack|unpack" -TimeoutMultiplier 2`

## Phase 7: Specs, Docs, Localization, and Closeout

### Required spec updates

Update `Specs/UI/UI_CommandMenuKeyboard.md`:

- Remove `[td]` markers for implemented commands.
- Document exact behavior for every command implemented by this plan.
- Preserve and verify existing `cmd/app/viewWidth` documentation as splitter view-width adjust behavior.
- Add `cmd/pane/alternateEdit`.
- Add `cmd/app/theme/selectNext` and `cmd/app/theme/selectPrev`.
- Document disabled-state behavior and shortcut fallback behavior.
- Document parameterized command id forms.
- Document dynamic menu population rules.

Update `Specs/Core/Core_SettingsStore.md`:

- Add new schema version.
- Document shared external-action settings and macro grammar for View, Alternate View, Edit, Alternate Edit, View With, Edit With, and User Menu.
- Document editors settings.
- Document user menu settings.
- Document pane view option settings.
- Document Make File List saved options.
- Document command-line history settings, if added.
- Document migration defaults.

Update `Specs/FileSystem/FileSystem_FileOperations.md`:

- Document pack/unpack operation semantics.
- Document cancellation and partial-output cleanup.
- Document archive overwrite policy.

Update `Specs/Testing/Testing_TestCoverage.md`:

- Add the new command selftest families and cases.
- Document archive/file-list/share/thumbnail/preview protected scenarios.

Update `Specs/Testing/Testing_PerformanceValidation.md` only if new metric names or scenario categories are added.

### Required docs updates

Update:

- `Docs/Todo.md`: remove completed command entries or move them to a completed section.
- `Docs/Preferences.md`: describe Editors and User Menu preferences.
- Any user-facing docs for file operations, view options, and command-line behavior.

### Localization

For English and French resource files:

- Remove `[todo]` labels from implemented menu entries.
- Add strings for:
  - no alternate viewer available
  - no alternate editor available
  - editor launch failure
  - invalid edit-new file name
  - editor combo labels for Edit New
  - unsupported plugin file-system command
  - alternate data stream removal and report labels
  - shell security failure
  - shortcut target resolution failure
  - reparse point, junction, and mount point target labels
  - command-line launch failure
  - empty user menu
  - pack/unpack validation and progress labels
  - make file list JSON, text macro, and saved option labels
  - shared directories labels
  - next theme and previous theme labels

### Final validation

Run:

- `.\build.ps1 -ProjectName RedSalamander`
- `.\build.ps1 -ProjectName RedSalamander -Configuration Release`
- `.\Tools\Run-AllTests.ps1 -Suite Commands -TimeoutMultiplier 2`
- `.\Tools\Run-AllTests.ps1 -Suite FileOps -TimeoutMultiplier 2`
- Any project-wide selftest command required by `Specs/Testing/Testing_SelfTests.md`

Before completion:

- Confirm no implemented command still has `[td]` in `Specs/UI/UI_CommandMenuKeyboard.md`.
- Confirm no implemented menu label still has `[todo]` in English or French resources.
- Confirm `cmd/app/viewWidth` remains the documented splitter view-width adjust command and keeps its existing shortcut/selftest coverage.
- Confirm no new user-facing strings are hardcoded in C++.
- Confirm new settings have migration defaults and roundtrip tests.
- Confirm perf evidence exists for thumbnails, preview, make file list, pack, unpack, and reread associations.
- Move this file from `Specs/Plans/WIP/` to `Specs/Plans/Done/`.

## Implementation Order and Merge Strategy

Use small reviewable changes:

1. Baseline inventory and stale-marker cleanup.
2. Command routing foundation and theme next/previous commands.
3. Viewer/editor settings and commands.
4. Shell/item utility commands.
5. Clipboard and command-line commands.
6. Pane view options.
7. List/share/association commands.
8. Archive pack/unpack commands.
9. Final docs, specs, localization, and perf evidence.

Do not start archive pack/unpack before the command routing foundation and file-operation selftest surfaces are green.

Do not remove a `[todo]` menu label because a command has a stub. Remove it only after the command's expected behavior is implemented, localized, and covered by deterministic tests.

## Risks and Mitigations

- Archive write support may require extending `Plugins/FileSystem7z`.
  - Mitigation: implement an archive operation abstraction first, then add `.zip` support as the minimum shippable format.
- Thumbnail and preview work can regress rendering responsiveness.
  - Mitigation: add instrumentation and bounded async work before enabling the UI command.
- Shell integration can leak COM resources if implemented manually.
  - Mitigation: use `wil::com_ptr`, WIL registry/handle wrappers, and focused selftests with shell dry-run providers.
- Command routing can diverge between shortcuts and menus.
  - Mitigation: make both paths call the same execution function and test both inputs for every new command.
- Settings schema growth can break older profiles.
  - Mitigation: add migration defaults and parse/write roundtrip tests before UI code depends on new settings.
- Multi-pane view state can drift between left/right commands and active-pane shortcuts.
  - Mitigation: centralize pane target resolution and test active-pane plus explicit left/right paths.

## Done Criteria

- Every command in the backlog inventory is implemented, context-disabled, or explicitly documented as unsupported for a specific file-system capability.
- Shortcut dispatch, menu dispatch, dynamic menu dispatch, and command registry metadata use the same command semantics.
- `cmd/app/viewWidth` is unchanged as the splitter view-width adjust command and remains documented/tested.
- `cmd/pane/editWidth` is no longer exposed as a missing command.
- English and French menus no longer mark implemented behavior as `[todo]`.
- `Specs/UI/UI_CommandMenuKeyboard.md` no longer marks implemented behavior as `[td]`.
- New durable behavior is documented in authoritative specs.
- Deterministic command selftests cover every command family.
- Perf evidence is archived for all responsive or file-operation-sensitive scenarios.
- The final implementation builds in Debug and Release.
- This plan is moved to `Specs/Plans/Done/`.

void RunPreferencesCommandsSelfTestCases(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite) noexcept
{
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_category_tree_uses_dxui_host_without_visible_legacy_treeview", [=](CaseState& state) noexcept {
        return TestPreferencesDialogCategoryTreeUsesDxUiHost(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_opens_with_french_satellite_resources", [=](CaseState& state) noexcept {
        return TestPreferencesDialogOpensWithFrenchSatelliteResources(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_category_tree_exposes_live_uia_selection", [=](CaseState& state) noexcept {
        return TestPreferencesDialogCategoryTreeExposesLiveUiaSelection(mainWindow, state);
    });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_preferences_dialog_shell_uses_dxui_header_footer_without_visible_legacy_shell_controls",
                      [=](CaseState& state) noexcept { return TestPreferencesDialogShellUsesDxUiChrome(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_page_host_uses_dxui_surface", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPageHostUsesDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_page_uses_dxui_combo_and_button_chrome", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersPageUsesDxUiChrome(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_page_uses_dxui_shell_chrome", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsPageUsesDxUiShellChrome(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_roundtrip_restores_dxui_surface", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsRoundTripRestoresDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_theme_cycle_keeps_surface_legible", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsThemeCycleKeepsSurfaceLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_roundtrip_restores_dxui_surface", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersRoundTripRestoresDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_theme_cycle_keeps_surface_legible", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersThemeCycleKeepsSurfaceLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_search_roundtrip_preserves_retained_state", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersSearchRoundTripPreservesRetainedState(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_search_action_updates_dxui_surface", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersSearchActionUpdatesDxSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_live_search_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersLiveSearchDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_remove_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersRemoveLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_reset_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersResetLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_add_update_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersAddUpdateLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_selection_survives_legacy_list_clear", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersSelectionSurvivesLegacyListClear(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugin_tree_selection_keeps_single_visible_pane", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginTreeSelectionKeepsSingleVisiblePane(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugin_tree_left_right_navigation_stays_on_dx_tree_path", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginTreeLeftRightNavigationStaysOnDxTreePath(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_search_roundtrip_preserves_retained_state", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsSearchRoundTripPreservesRetainedState(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_preferences_dialog_plugins_custom_paths_selection_roundtrip_preserves_retained_state", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsCustomPathsSelectionRoundTripPreservesRetainedState(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_main_selection_survives_legacy_list_clear", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsMainSelectionSurvivesLegacyListClear(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_custom_paths_selection_survives_legacy_list_clear", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsCustomPathsSelectionSurvivesLegacyListClear(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_main_checkbox_survives_legacy_row_state_change", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsMainCheckboxSurvivesLegacyRowStateChange(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_main_checkbox_space_toggles_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsMainCheckboxSpaceTogglesLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_main_checkbox_click_toggles_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsMainCheckboxClickTogglesLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_main_list_header_drag_reorders_columns_without_sort", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsMainListHeaderDragReordersColumnsWithoutSort(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_main_list_header_resize_changes_visible_width", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsMainListHeaderResizeChangesVisibleWidth(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_reordered_resized_columns_survive_search_roundtrip", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsReorderedResizedColumnsSurviveSearchRoundTrip(mainWindow, state);
    });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_preferences_dialog_plugins_reordered_resized_copy_follows_visible_columns_after_search_roundtrip",
                      [=](CaseState& state) noexcept
    { return TestPreferencesDialogPluginsReorderedResizedCopyFollowsVisibleColumnsAfterSearchRoundTrip(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"cmd_preferences_dialog_plugins_reordered_resized_columns_survive_sort_cycles_and_search_roundtrip", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsReorderedResizedColumnsSurviveSortCyclesAndSearchRoundTrip(mainWindow, state);
    });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_preferences_dialog_plugins_reordered_resized_copy_follows_visible_columns_after_sort_cycles_and_search_roundtrip",
                      [=](CaseState& state) noexcept
    { return TestPreferencesDialogPluginsReorderedResizedCopyFollowsVisibleColumnsAfterSortCyclesAndSearchRoundTrip(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_page_exposes_live_uia_grid_selection", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsPageExposesLiveGridSelection(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_custom_paths_page_exposes_live_uia_grid_selection", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsCustomPathsPageExposesLiveGridSelection(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_long_run_list_scrolling_stays_bounded", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsLongRunListScrollingStaysBounded(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_custom_paths_long_run_list_scrolling_stays_bounded", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsCustomPathsLongRunListScrollingStaysBounded(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_long_run_list_scrolling_stays_bounded", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardLongRunListScrollingStaysBounded(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_long_run_list_scrolling_stays_bounded", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersLongRunListScrollingStaysBounded(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_long_run_list_scrolling_stays_bounded", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesLongRunListScrollingStaysBounded(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_page_uses_dxui_shell_chrome", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesPageUsesDxUiShellChrome(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_roundtrip_restores_dxui_surface", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesRoundTripRestoresDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_theme_cycle_keeps_surface_legible", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesThemeCycleKeepsSurfaceLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_page_exposes_live_uia_grid_selection", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersPageExposesLiveGridSelection(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_tab_traversal_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersTabTraversalLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_pointer_click_selects_live_dx_row", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersPointerClickSelectsLiveDxRow(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_header_drag_reorders_columns_without_sort", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersHeaderDragReordersColumnsWithoutSort(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_copy_follows_reordered_columns", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersCopyFollowsReorderedColumns(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_header_resize_changes_visible_width", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersHeaderResizeChangesVisibleWidth(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_header_resize_survives_search_roundtrip", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersHeaderResizeSurvivesSearchRoundTrip(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_reordered_columns_survive_search_roundtrip", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersReorderedColumnsSurviveSearchRoundTrip(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_preferences_dialog_viewers_reordered_copy_follows_visible_columns_after_search_roundtrip", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersReorderedCopyFollowsVisibleColumnsAfterSearchRoundTrip(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_reordered_resized_columns_survive_search_roundtrip", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersReorderedResizedColumnsSurviveSearchRoundTrip(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_viewers_reordered_resized_columns_survive_sort_cycles", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersReorderedResizedColumnsSurviveSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_preferences_dialog_viewers_reordered_resized_copy_follows_visible_columns_after_sort_cycles", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersReorderedResizedCopyFollowsVisibleColumnsAfterSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_preferences_dialog_viewers_reordered_resized_columns_survive_sort_cycles_and_search_roundtrip", [=](CaseState& state) noexcept {
        return TestPreferencesDialogViewersReorderedResizedColumnsSurviveSortCyclesAndSearchRoundTrip(mainWindow, state);
    });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_preferences_dialog_viewers_reordered_resized_copy_follows_visible_columns_after_sort_cycles_and_search_roundtrip",
                      [=](CaseState& state) noexcept
    { return TestPreferencesDialogViewersReorderedResizedCopyFollowsVisibleColumnsAfterSortCyclesAndSearchRoundTrip(mainWindow, state); });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_preferences_dialog_viewers_reordered_resized_copy_follows_visible_columns_after_search_roundtrip",
                      [=](CaseState& state) noexcept
    { return TestPreferencesDialogViewersReorderedResizedCopyFollowsVisibleColumnsAfterSearchRoundTrip(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_page_exposes_live_uia_grid_selection", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardPageExposesLiveGridSelection(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_header_drag_reorders_columns_without_sort", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardHeaderDragReordersColumnsWithoutSort(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_header_resize_changes_visible_width", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardHeaderResizeChangesVisibleWidth(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_reordered_resized_columns_stay_stable", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardReorderedResizedColumnsStayStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_reordered_resized_columns_survive_search_roundtrip", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardReorderedResizedColumnsSurviveSearchRoundTrip(mainWindow, state);
    });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_preferences_dialog_keyboard_reordered_resized_copy_follows_visible_columns_after_search_roundtrip",
                      [=](CaseState& state) noexcept
    { return TestPreferencesDialogKeyboardReorderedResizedCopyFollowsVisibleColumnsAfterSearchRoundTrip(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_copy_follows_reordered_columns", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardCopyFollowsReorderedColumns(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_search_roundtrip_preserves_retained_state", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesSearchRoundTripPreservesRetainedState(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_selection_survives_legacy_combo_clear", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesSelectionSurvivesLegacyComboClear(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_color_selection_survives_legacy_list_clear", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesColorSelectionSurvivesLegacyListClear(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_page_exposes_live_uia_grid_selection", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesPageExposesLiveGridSelection(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_tab_traversal_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesTabTraversalLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_pointer_click_selects_live_dx_row", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesPointerClickSelectsLiveDxRow(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_header_drag_reorders_columns_without_sort", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesHeaderDragReordersColumnsWithoutSort(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_header_resize_changes_visible_width", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesHeaderResizeChangesVisibleWidth(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_reordered_resized_columns_survive_search_roundtrip", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesReorderedResizedColumnsSurviveSearchRoundTrip(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_preferences_dialog_themes_reordered_resized_copy_follows_visible_columns_after_search_roundtrip", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesReorderedResizedCopyFollowsVisibleColumnsAfterSearchRoundTrip(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_preferences_dialog_themes_reordered_resized_columns_survive_sort_cycles_and_search_roundtrip", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesReorderedResizedColumnsSurviveSortCyclesAndSearchRoundTrip(mainWindow, state);
    });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_preferences_dialog_themes_reordered_resized_copy_follows_visible_columns_after_sort_cycles_and_search_roundtrip",
                      [=](CaseState& state) noexcept
    { return TestPreferencesDialogThemesReorderedResizedCopyFollowsVisibleColumnsAfterSortCyclesAndSearchRoundTrip(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_copy_follows_reordered_columns", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesCopyFollowsReorderedColumns(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_general_page_uses_dxui_toggle_cards", [=](CaseState& state) noexcept {
        return TestPreferencesDialogGeneralPageUsesDxUiToggleCards(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_general_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogGeneralLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_general_dxui_customization_preview_and_cancel", [=](CaseState& state) noexcept {
        return TestPreferencesDialogGeneralDxUiCustomizationPreviewAndCancel(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_general_window_backdrop_apply_updates_supported_windows", [=](CaseState& state) noexcept {
        return TestPreferencesDialogGeneralWindowBackdropApplyUpdatesSupportedWindows(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_general_tab_traversal_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogGeneralTabTraversalLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_general_roundtrip_restores_dxui_surface", [=](CaseState& state) noexcept {
        return TestPreferencesDialogGeneralRoundTripRestoresDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_general_theme_cycle_keeps_surface_legible", [=](CaseState& state) noexcept {
        return TestPreferencesDialogGeneralThemeCycleKeepsSurfaceLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_panes_page_uses_dxui_statics_and_toggles", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPanesPageUsesDxUiStaticsAndToggles(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_shell_footer_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogShellFooterLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_shell_footer_access_keys_route_expected_actions", [=](CaseState& state) noexcept {
        return TestPreferencesDialogShellFooterAccessKeysRouteExpectedActions(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_panes_roundtrip_restores_dxui_surface", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPanesRoundTripRestoresDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_panes_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPanesLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_panes_theme_cycle_keeps_surface_legible", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPanesThemeCycleKeepsSurfaceLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_panes_tab_traversal_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPanesTabTraversalLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_panes_history_size_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPanesHistorySizeLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_panes_combo_then_toggle_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPanesComboThenToggleLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_hot_paths_page_uses_dxui_statics_and_toggles", [=](CaseState& state) noexcept {
        return TestPreferencesDialogHotPathsPageUsesDxUiStaticsAndToggles(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_hot_paths_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogHotPathsLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_hot_paths_open_prefs_toggle_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogHotPathsOpenPrefsToggleLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_hot_paths_browse_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogHotPathsBrowseLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_hot_paths_tab_traversal_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogHotPathsTabTraversalLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_hot_paths_roundtrip_restores_dxui_surface", [=](CaseState& state) noexcept {
        return TestPreferencesDialogHotPathsRoundTripRestoresDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_hot_paths_theme_cycle_keeps_surface_legible", [=](CaseState& state) noexcept {
        return TestPreferencesDialogHotPathsThemeCycleKeepsSurfaceLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_theme_cycle_keeps_surface_legible", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardThemeCycleKeepsSurfaceLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_page_uses_dxui_shell_chrome", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardPageUsesDxUiShellChrome(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_roundtrip_restores_dxui_surface", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardRoundTripRestoresDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_tab_traversal_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardTabTraversalLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_search_roundtrip_preserves_retained_state", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardSearchRoundTripPreservesRetainedState(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_search_action_updates_dxui_surface", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardSearchActionUpdatesDxSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_live_search_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardLiveSearchDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_reset_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardResetLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_export_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardExportLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_import_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardImportLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_remove_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardRemoveLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_assign_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardAssignLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_commit_assign_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardCommitAssignLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_capture_preview_and_assign_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardCapturePreviewAndAssignLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_replace_assign_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardReplaceAssignLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_keyboard_swap_assign_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogKeyboardSwapAssignLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_search_action_updates_dxui_surface", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsSearchActionUpdatesDxSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_live_search_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsLiveSearchDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_tab_traversal_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsTabTraversalLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_empty_custom_paths_placeholder", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsEmptyCustomPathsPlaceholder(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_custom_paths_remove_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsCustomPathsRemoveLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_custom_paths_add_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsCustomPathsAddLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_configure_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsConfigureLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_test_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsTestLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_plugins_test_all_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogPluginsTestAllLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_search_action_updates_dxui_surface", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesSearchActionUpdatesDxSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_live_search_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesLiveSearchDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_reset_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesResetLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_duplicate_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesDuplicateLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_clear_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesClearLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_set_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesSetLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_apply_temporarily_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesApplyTemporarilyLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_save_theme_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesSaveThemeLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_themes_load_from_file_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogThemesLoadFromFileLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_advanced_page_uses_dxui_statics_and_toggles", [=](CaseState& state) noexcept {
        return TestPreferencesDialogAdvancedPageUsesDxUiStaticsAndToggles(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_advanced_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogAdvancedLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_advanced_theme_cycle_keeps_surface_legible", [=](CaseState& state) noexcept {
        return TestPreferencesDialogAdvancedThemeCycleKeepsSurfaceLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_advanced_filter_preset_custom_mask_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogAdvancedFilterPresetCustomMaskLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_advanced_tab_traversal_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogAdvancedTabTraversalLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_advanced_roundtrip_restores_dxui_surface", [=](CaseState& state) noexcept {
        return TestPreferencesDialogAdvancedRoundTripRestoresDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_editors_and_mouse_pages_use_dxui_statics", [=](CaseState& state) noexcept {
        return TestPreferencesDialogEditorsAndMousePagesUseDxUiStatics(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_editors_mouse_roundtrip_restore_dxui_notes", [=](CaseState& state) noexcept {
        return TestPreferencesDialogEditorsAndMouseRoundTripRestoreDxUiNotes(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_file_operations_page_uses_dxui_controls", [=](CaseState& state) noexcept {
        return TestPreferencesDialogFileOperationsPageUsesDxUiControls(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_file_operations_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogFileOperationsLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_file_operations_custom_bandwidth_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogFileOperationsCustomBandwidthLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_file_operations_theme_cycle_keeps_surface_legible", [=](CaseState& state) noexcept {
        return TestPreferencesDialogFileOperationsThemeCycleKeepsSurfaceLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_file_operations_tab_traversal_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogFileOperationsTabTraversalLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_file_operations_roundtrip_restores_dxui_surface", [=](CaseState& state) noexcept {
        return TestPreferencesDialogFileOperationsRoundTripRestoresDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_compare_directories_page_uses_dxui_statics", [=](CaseState& state) noexcept {
        return TestPreferencesDialogCompareDirectoriesPageUsesDxUiStatics(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_compare_directories_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogCompareDirectoriesLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_preferences_dialog_compare_directories_content_workers_ignore_files_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogCompareDirectoriesContentWorkersIgnoreFilesLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_compare_directories_tail_toggles_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogCompareDirectoriesTailTogglesLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_compare_directories_tab_traversal_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPreferencesDialogCompareDirectoriesTabTraversalLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_compare_directories_roundtrip_restores_dxui_surface", [=](CaseState& state) noexcept {
        return TestPreferencesDialogCompareDirectoriesRoundTripRestoresDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_compare_directories_theme_cycle_keeps_surface_legible", [=](CaseState& state) noexcept {
        return TestPreferencesDialogCompareDirectoriesThemeCycleKeepsSurfaceLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_category_tree_keyboard_navigation_updates_category", [=](CaseState& state) noexcept {
        return TestPreferencesDialogCategoryTreeKeyboardNavigation(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_editors_mouse_live_dx_notes", [=](CaseState& state) noexcept {
        return TestPreferencesDialogEditorsAndMouseLiveDxNotes(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_editors_mouse_tab_skips_note_surface", [=](CaseState& state) noexcept {
        return TestPreferencesDialogEditorsAndMouseTabSkipNoteSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation", [=](CaseState& state) noexcept {
        return TestPreferencesDialogCategoryTreeReverseKeyboardNavigation(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_category_tree_page_navigation_stays_on_dx_tree_path", [=](CaseState& state) noexcept {
        return TestPreferencesDialogCategoryTreePageNavigation(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_category_tree_accepts_wheel_scrolling_without_selection_churn", [=](CaseState& state) noexcept {
        return TestPreferencesDialogCategoryTreeWheelScrollingWorks(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_category_tree_keyboard_expand_collapse_and_child_entry", [=](CaseState& state) noexcept {
        return TestPreferencesDialogCategoryTreeKeyboardExpandCollapseAndChildEntry(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_category_tree_boundary_navigation_from_scrolled_state", [=](CaseState& state) noexcept {
        return TestPreferencesDialogCategoryTreeBoundaryNavigationFromScrolledState(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_category_switches_do_not_churn_tree_host", [=](CaseState& state) noexcept {
        return TestPreferencesDialogCategorySwitchesDoNotChurnTreeHost(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_scroll_host_preserves_retained_page_state", [=](CaseState& state) noexcept {
        return TestPreferencesDialogScrollHostPreservesRetainedPageState(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_preferences_dialog_rapid_switches_keep_page_specific_uia_subtrees", [=](CaseState& state) noexcept {
        return TestPreferencesDialogRapidSwitchesKeepPageSpecificUiaSubtrees(mainWindow, state);
    });
}

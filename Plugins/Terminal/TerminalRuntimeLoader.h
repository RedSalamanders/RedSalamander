#pragma once

#include <string>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <bcrypt.h>
#include <wincodec.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

// Ghostty declarations are private implementation detail of Terminal.dll.
#include <ghostty/vt.h>

class TerminalPngDecodeThreadContext final
{
public:
    TerminalPngDecodeThreadContext() noexcept;
    ~TerminalPngDecodeThreadContext() noexcept;

    TerminalPngDecodeThreadContext(const TerminalPngDecodeThreadContext&) = delete;
    TerminalPngDecodeThreadContext(TerminalPngDecodeThreadContext&&) = delete;
    TerminalPngDecodeThreadContext& operator=(const TerminalPngDecodeThreadContext&) = delete;
    TerminalPngDecodeThreadContext& operator=(TerminalPngDecodeThreadContext&&) = delete;

    [[nodiscard]] HRESULT Status() const noexcept;

private:
    wil::unique_couninitialize_call _coUninitialize{false};
    wil::com_ptr<IWICImagingFactory> _factory;
    IWICImagingFactory* _previousFactory = nullptr; // Non-owning thread-local context restored on scope exit.
    bool _installed = false;
    HRESULT _status = E_UNEXPECTED;
};

class TerminalRuntimeLoader final
{
public:
    TerminalRuntimeLoader() = default;
    ~TerminalRuntimeLoader() = default;

    TerminalRuntimeLoader(const TerminalRuntimeLoader&) = delete;
    TerminalRuntimeLoader(TerminalRuntimeLoader&&) = delete;
    TerminalRuntimeLoader& operator=(const TerminalRuntimeLoader&) = delete;
    TerminalRuntimeLoader& operator=(TerminalRuntimeLoader&&) = delete;

    [[nodiscard]] HRESULT Load() noexcept;
    void Unload() noexcept;
    [[nodiscard]] bool IsLoaded() const noexcept;
    [[nodiscard]] const std::wstring& ErrorText() const noexcept;
    [[nodiscard]] GhosttyResult GetTerminalMode(GhosttyTerminal terminal, GhosttyMode mode, bool& value) const noexcept;

    decltype(&ghostty_build_info) buildInfo = nullptr;
    decltype(&ghostty_terminal_new) terminalNew = nullptr;
    decltype(&ghostty_terminal_free) terminalFree = nullptr;
    decltype(&ghostty_terminal_resize) terminalResize = nullptr;
    decltype(&ghostty_terminal_set) terminalSet = nullptr;
    decltype(&ghostty_terminal_get) terminalGet = nullptr;
    decltype(&ghostty_terminal_grid_ref) terminalGridRef = nullptr;
    decltype(&ghostty_grid_ref_hyperlink_uri) gridRefHyperlinkUri = nullptr;
    decltype(&ghostty_terminal_grid_ref_track) terminalGridRefTrack = nullptr;
    decltype(&ghostty_tracked_grid_ref_free) trackedGridRefFree = nullptr;
    decltype(&ghostty_tracked_grid_ref_snapshot) trackedGridRefSnapshot = nullptr;
    decltype(&ghostty_terminal_vt_write) terminalVtWrite = nullptr;
    decltype(&ghostty_terminal_select_word) terminalSelectWord = nullptr;
    decltype(&ghostty_terminal_select_all) terminalSelectAll = nullptr;
    decltype(&ghostty_terminal_selection_format_buf) terminalSelectionFormatBuffer = nullptr;
    decltype(&ghostty_terminal_scroll_viewport) terminalScrollViewport = nullptr;
    decltype(&ghostty_paste_is_safe) pasteIsSafe = nullptr;
    decltype(&ghostty_paste_encode) pasteEncode = nullptr;
    decltype(&ghostty_render_state_new) renderStateNew = nullptr;
    decltype(&ghostty_render_state_free) renderStateFree = nullptr;
    decltype(&ghostty_render_state_begin_update) renderStateBeginUpdate = nullptr;
    decltype(&ghostty_render_state_end_update) renderStateEndUpdate = nullptr;
    decltype(&ghostty_render_state_get) renderStateGet = nullptr;
    decltype(&ghostty_render_state_set) renderStateSet = nullptr;
    decltype(&ghostty_render_state_row_iterator_new) renderRowIteratorNew = nullptr;
    decltype(&ghostty_render_state_row_iterator_free) renderRowIteratorFree = nullptr;
    decltype(&ghostty_render_state_row_iterator_next) renderRowIteratorNext = nullptr;
    decltype(&ghostty_render_state_row_get) renderRowGet = nullptr;
    decltype(&ghostty_render_state_row_set) renderRowSet = nullptr;
    decltype(&ghostty_render_state_row_cells_new) renderRowCellsNew = nullptr;
    decltype(&ghostty_render_state_row_cells_free) renderRowCellsFree = nullptr;
    decltype(&ghostty_render_state_row_cells_next) renderRowCellsNext = nullptr;
    decltype(&ghostty_render_state_row_cells_get) renderRowCellsGet = nullptr;
    decltype(&ghostty_cell_get) cellGet = nullptr;
    decltype(&ghostty_key_encoder_new) keyEncoderNew = nullptr;
    decltype(&ghostty_key_encoder_free) keyEncoderFree = nullptr;
    decltype(&ghostty_key_encoder_setopt_from_terminal) keyEncoderSetFromTerminal = nullptr;
    decltype(&ghostty_key_encoder_encode) keyEncoderEncode = nullptr;
    decltype(&ghostty_key_event_new) keyEventNew = nullptr;
    decltype(&ghostty_key_event_free) keyEventFree = nullptr;
    decltype(&ghostty_key_event_set_action) keyEventSetAction = nullptr;
    decltype(&ghostty_key_event_set_key) keyEventSetKey = nullptr;
    decltype(&ghostty_key_event_set_mods) keyEventSetMods = nullptr;
    decltype(&ghostty_key_event_set_utf8) keyEventSetUtf8 = nullptr;
    decltype(&ghostty_key_event_set_unshifted_codepoint) keyEventSetUnshiftedCodepoint = nullptr;
    decltype(&ghostty_mouse_encoder_new) mouseEncoderNew = nullptr;
    decltype(&ghostty_mouse_encoder_free) mouseEncoderFree = nullptr;
    decltype(&ghostty_mouse_encoder_setopt) mouseEncoderSetOption = nullptr;
    decltype(&ghostty_mouse_encoder_setopt_from_terminal) mouseEncoderSetFromTerminal = nullptr;
    decltype(&ghostty_mouse_encoder_encode) mouseEncoderEncode = nullptr;
    decltype(&ghostty_mouse_event_new) mouseEventNew = nullptr;
    decltype(&ghostty_mouse_event_free) mouseEventFree = nullptr;
    decltype(&ghostty_mouse_event_set_action) mouseEventSetAction = nullptr;
    decltype(&ghostty_mouse_event_set_button) mouseEventSetButton = nullptr;
    decltype(&ghostty_mouse_event_clear_button) mouseEventClearButton = nullptr;
    decltype(&ghostty_mouse_event_set_mods) mouseEventSetMods = nullptr;
    decltype(&ghostty_mouse_event_set_position) mouseEventSetPosition = nullptr;
    decltype(&ghostty_alloc) alloc = nullptr;
    decltype(&ghostty_free) free = nullptr;
    decltype(&ghostty_sys_set) sysSet = nullptr;
    decltype(&ghostty_kitty_graphics_get) kittyGraphicsGet = nullptr;
    decltype(&ghostty_kitty_graphics_image) kittyGraphicsImage = nullptr;
    decltype(&ghostty_kitty_graphics_image_get_multi) kittyGraphicsImageGetMulti = nullptr;
    decltype(&ghostty_kitty_graphics_placement_iterator_new) kittyPlacementIteratorNew = nullptr;
    decltype(&ghostty_kitty_graphics_placement_iterator_free) kittyPlacementIteratorFree = nullptr;
    decltype(&ghostty_kitty_graphics_placement_next) kittyPlacementNext = nullptr;
    decltype(&ghostty_kitty_graphics_placement_get) kittyPlacementGet = nullptr;
    decltype(&ghostty_kitty_graphics_placement_render_info) kittyPlacementRenderInfo = nullptr;
    decltype(&ghostty_formatter_terminal_new) formatterTerminalNew = nullptr;
    decltype(&ghostty_formatter_format_buf) formatterFormatBuffer = nullptr;
    decltype(&ghostty_formatter_free) formatterFree = nullptr;

private:
    // Keep the verified image open without write/delete sharing for the full module lifetime.
    // Member order is intentional: destruction releases the module before this file handle.
    wil::unique_hfile _runtimeFile;
    wil::unique_hmodule _module;
    std::wstring _errorText;
    bool _systemCallbacksAcquired = false;
};

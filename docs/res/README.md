# Documentation Image Inventory

Publish application screenshots from deterministic native test-harness scenarios.
Use realistic invented identifying data, inspect the complete frame and retain
source/build/scenario/theme/size/DPI/hash provenance. Seed data before capture;
never use desktop grabs, blur, masks, redaction labels or later pixel edits.
Actual test I/O stays in a marked `X:\RedSalamander.Perf` even when the displayed
example path is `C:\Demo\Media\Workshop\session-01.mp4`.

The authoritative [screenshot contract](../../Specs/Testing/Testing_SelfTests.md#documentation-screenshot-scenarios)
and [capture workflow](../dev/TestingAndScreenshots.md) apply to all new/replaced images.

## Current File Operations captures

The ten public frames come from the private English `FileOps_ReviewRemediationContracts`
fixture with invented Workshop paths. No operation is admitted on those displayed
paths, and an empty owned breadcrumb store keeps real history out of the frames.
The test asserts exact accessible identities, provisional/final counts, multiple
unclipped note lines above progress, and a selected issue before saving native frames.
Each PNG has an adjacent `.provenance.json`; the documentation drift guard checks
the published hash and review metadata. Human whole-frame privacy review remains required.

| Image | Scenario |
| --- | --- |
| [file-operations-popup.png](file-operations-popup.png) | Running transfer with final totals, counts, Throughput speed and automatic budget. [Provenance](file-operations-popup.provenance.json). |
| [file-operations-counting.png](file-operations-counting.png) | Open discovery with provisional counts and the Counting label. [Provenance](file-operations-counting.provenance.json). |
| [file-operations-move-totals.png](file-operations-move-totals.png) | Move with plain recursive totals. [Provenance](file-operations-move-totals.provenance.json). |
| [file-operations-file-progress.png](file-operations-file-progress.png) | Copy with folder totals and file completion fractions. [Provenance](file-operations-file-progress.provenance.json). |
| [file-operations-paused.png](file-operations-paused.png) | Paused Copy retaining progress. [Provenance](file-operations-paused.provenance.json). |
| [file-operations-queued.png](file-operations-queued.png) | Copy waiting for admission. [Provenance](file-operations-queued.provenance.json). |
| [file-operations-completed.png](file-operations-completed.png) | Clean terminal Copy result. [Provenance](file-operations-completed.provenance.json). |
| [file-operations-conflict.png](file-operations-conflict.png) | Destination-exists decision with invented metadata. [Provenance](file-operations-conflict.provenance.json). |
| [file-operations-wrapped-note.png](file-operations-wrapped-note.png) | Failed Move with a complete, genuinely wrapped terminal note. [Provenance](file-operations-wrapped-note.provenance.json). |
| [file-operations-issues.png](file-operations-issues.png) | Actual Issues pane with one selected source-missing issue. [Provenance](file-operations-issues.provenance.json). |

## Product scenario captures

The [Visual Tour](../VisualTour.md) links 45 additional native frames: 26 workspace,
Search, Preferences, plugin, connection and palette scenes; eight actual viewer
examples; and thirteen Space lifecycle/workspace states. The [gallery index](product-scenarios.provenance.json)
binds every image to adjacent build/scenario/DPI/hash metadata. Captures use private
default settings, generated content, fixed synthetic metadata and the existing fake
MTP picker. No real search, account connection or device enumeration is admitted.
The VLC example demonstrates missing-runtime setup guidance, not playback.
Workspace examples show a complete native FolderWindow component in a test host;
Space examples show its complete native viewer surface. Scrollable controls retain
their real scrollbars, and the guides describe options below the visible viewport.

### Workspace and configuration

| Image | Scenario |
| --- | --- |
| [folder-view-brief-detailed.png](folder-view-brief-detailed.png) | Folder view brief detailed. [Provenance](folder-view-brief-detailed.provenance.json). |
| [folder-view-extra-detailed.png](folder-view-extra-detailed.png) | Folder view extra detailed. [Provenance](folder-view-extra-detailed.provenance.json). |
| [theme-light.png](theme-light.png) | Theme light. [Provenance](theme-light.provenance.json). |
| [theme-dark.png](theme-dark.png) | Theme dark. [Provenance](theme-dark.provenance.json). |
| [theme-high-contrast-app.png](theme-high-contrast-app.png) | Theme high contrast app. [Provenance](theme-high-contrast-app.provenance.json). |
| [theme-rainbow.png](theme-rainbow.png) | Theme rainbow. [Provenance](theme-rainbow.provenance.json). |
| [preferences-general.png](preferences-general.png) | Preferences general. [Provenance](preferences-general.provenance.json). |
| [preferences-panes.png](preferences-panes.png) | Preferences panes. [Provenance](preferences-panes.provenance.json). |
| [preferences-viewers.png](preferences-viewers.png) | Preferences viewers. [Provenance](preferences-viewers.provenance.json). |
| [preferences-editors.png](preferences-editors.png) | Preferences editors. [Provenance](preferences-editors.provenance.json). |
| [preferences-keyboard.png](preferences-keyboard.png) | Preferences keyboard. [Provenance](preferences-keyboard.provenance.json). |
| [preferences-mouse.png](preferences-mouse.png) | Preferences mouse. [Provenance](preferences-mouse.provenance.json). |
| [preferences-themes.png](preferences-themes.png) | Preferences themes. [Provenance](preferences-themes.provenance.json). |
| [plugins.png](plugins.png) | Plugins. [Provenance](plugins.provenance.json). |
| [preferences-advanced.png](preferences-advanced.png) | Preferences advanced. [Provenance](preferences-advanced.provenance.json). |
| [preferences-compare-directories.png](preferences-compare-directories.png) | Preferences compare directories. [Provenance](preferences-compare-directories.provenance.json). |
| [preferences-hot-paths.png](preferences-hot-paths.png) | Preferences hot paths. [Provenance](preferences-hot-paths.provenance.json). |
| [preferences-file-operations.png](preferences-file-operations.png) | Preferences file operations. [Provenance](preferences-file-operations.provenance.json). |
| [preferences-user-menu.png](preferences-user-menu.png) | Preferences user menu. [Provenance](preferences-user-menu.provenance.json). |
| [preferences-monitor.png](preferences-monitor.png) | Preferences monitor. [Provenance](preferences-monitor.provenance.json). |
| [preferences-plugin-mtp.png](preferences-plugin-mtp.png) | Preferences plugin mtp. [Provenance](preferences-plugin-mtp.provenance.json). |
| [find-files.png](find-files.png) | Find files. [Provenance](find-files.provenance.json). |
| [find-files-empty.png](find-files-empty.png) | Find files empty. [Provenance](find-files-empty.provenance.json). |
| [command-palette.png](command-palette.png) | Command palette. [Provenance](command-palette.provenance.json). |
| [connections.png](connections.png) | Connections. [Provenance](connections.provenance.json). |
| [connections-mtp.png](connections-mtp.png) | Connections mtp. [Provenance](connections-mtp.provenance.json). |

### Generated-content viewers

| Image | Scenario |
| --- | --- |
| [viewer-text.png](viewer-text.png) | Viewer text. [Provenance](viewer-text.provenance.json). |
| [viewer-text-hex.png](viewer-text-hex.png) | Viewer text hex. [Provenance](viewer-text-hex.provenance.json). |
| [viewer-imgraw.png](viewer-imgraw.png) | Viewer imgraw. [Provenance](viewer-imgraw.provenance.json). |
| [viewer-pe.png](viewer-pe.png) | Viewer pe. [Provenance](viewer-pe.provenance.json). |
| [viewer-web.png](viewer-web.png) | Viewer web. [Provenance](viewer-web.provenance.json). |
| [viewer-json.png](viewer-json.png) | Viewer json. [Provenance](viewer-json.provenance.json). |
| [viewer-markdown.png](viewer-markdown.png) | Viewer markdown. [Provenance](viewer-markdown.provenance.json). |
| [viewer-vlc-unavailable.png](viewer-vlc-unavailable.png) | Viewer vlc unavailable. [Provenance](viewer-vlc-unavailable.provenance.json). |

### Space lifecycle and layouts

| Image | Scenario |
| --- | --- |
| [viewer-space-scanning.png](viewer-space-scanning.png) | Viewer space scanning. [Provenance](viewer-space-scanning.provenance.json). |
| [viewer-space-inflight.png](viewer-space-inflight.png) | Viewer space inflight. [Provenance](viewer-space-inflight.provenance.json). |
| [viewer-space-analyst.png](viewer-space-analyst.png) | Viewer space analyst. [Provenance](viewer-space-analyst.provenance.json). |
| [viewer-space-refreshing.png](viewer-space-refreshing.png) | Viewer space refreshing. [Provenance](viewer-space-refreshing.provenance.json). |
| [viewer-space-refresh-canceled.png](viewer-space-refresh-canceled.png) | Viewer space refresh canceled. [Provenance](viewer-space-refresh-canceled.provenance.json). |
| [viewer-space-refresh-refused.png](viewer-space-refresh-refused.png) | Viewer space refresh refused. [Provenance](viewer-space-refresh-refused.provenance.json). |
| [viewer-space-empty.png](viewer-space-empty.png) | Viewer space empty. [Provenance](viewer-space-empty.provenance.json). |
| [viewer-space-error.png](viewer-space-error.png) | Viewer space error. [Provenance](viewer-space-error.provenance.json). |
| [viewer-space-explorer.png](viewer-space-explorer.png) | Viewer space explorer. [Provenance](viewer-space-explorer.provenance.json). |
| [viewer-space-bottom-panel.png](viewer-space-bottom-panel.png) | Viewer space bottom panel. [Provenance](viewer-space-bottom-panel.provenance.json). |
| [viewer-space-map-only.png](viewer-space-map-only.png) | Viewer space map only. [Provenance](viewer-space-map-only.provenance.json). |

## Terminal, utility, provider and playback captures

The [supplemental capture index](supplemental-scenarios.provenance.json) records
21 further whole native frames. Terminal uses the real engine with a fixed VT
transcript and no shell/PTY/profile/history. Utility forms cancel without applying.
FTP, Google Drive, OneDrive and S3 browse owned loopback endpoints; MTP browses its
RAM backend. Playback and pause use a real pinned VLC runtime with an original,
generated silent AVI. These examples use no personal accounts, devices or files.

| Image | Scenario |
| --- | --- |
| [change-case.png](change-case.png) | Change case. [Provenance](change-case.provenance.json). |
| [change-attributes.png](change-attributes.png) | Change attributes. [Provenance](change-attributes.provenance.json). |
| [selection-mask.png](selection-mask.png) | Selection mask. [Provenance](selection-mask.provenance.json). |
| [file-filter.png](file-filter.png) | File filter. [Provenance](file-filter.provenance.json). |
| [make-file-list.png](make-file-list.png) | Make file list. [Provenance](make-file-list.provenance.json). |
| [create-directory.png](create-directory.png) | Create directory. [Provenance](create-directory.provenance.json). |
| [new-file.png](new-file.png) | New file. [Provenance](new-file.provenance.json). |
| [archive-pack.png](archive-pack.png) | Archive pack. [Provenance](archive-pack.provenance.json). |
| [archive-unpack.png](archive-unpack.png) | Archive unpack. [Provenance](archive-unpack.provenance.json). |
| [shared-directories.png](shared-directories.png) | Shared directories. [Provenance](shared-directories.provenance.json). |
| [shortcuts-terminal.png](shortcuts-terminal.png) | Shortcuts terminal. [Provenance](shortcuts-terminal.provenance.json). |
| [terminal-embedded.png](terminal-embedded.png) | Terminal embedded. [Provenance](terminal-embedded.provenance.json). |
| [terminal-floating-tabs.png](terminal-floating-tabs.png) | Terminal floating tabs. [Provenance](terminal-floating-tabs.provenance.json). |
| [terminal-find.png](terminal-find.png) | Terminal find. [Provenance](terminal-find.provenance.json). |
| [provider-ftp.png](provider-ftp.png) | Provider ftp. [Provenance](provider-ftp.provenance.json). |
| [provider-google-drive.png](provider-google-drive.png) | Provider google drive. [Provenance](provider-google-drive.provenance.json). |
| [provider-onedrive.png](provider-onedrive.png) | Provider onedrive. [Provenance](provider-onedrive.provenance.json). |
| [provider-s3.png](provider-s3.png) | Provider s3. [Provenance](provider-s3.provenance.json). |
| [provider-mtp.png](provider-mtp.png) | Provider mtp. [Provenance](provider-mtp.provenance.json). |
| [viewer-vlc-playback.png](viewer-vlc-playback.png) | Viewer vlc playback. [Provenance](viewer-vlc-playback.provenance.json). |
| [viewer-vlc-paused.png](viewer-vlc-paused.png) | Viewer vlc paused. [Provenance](viewer-vlc-paused.provenance.json). |

## Retained illustrations

The earlier static illustrations below were privacy-screened for the October 19
documentation pass. They show generated sample content, ordinary product labels
or naturally empty fields. Their older capture history does not qualify the
current build or waive the harness/provenance requirements for future replacements.
Control galleries are native DxUi harness illustrations. The current Space lifecycle set has its own [capture index](viewer-space-captures.provenance.json).
`Banner.jpg` is branding artwork, not an application screenshot.

| Image | Subject |
| --- | --- |
| [Banner.jpg](Banner.jpg) | Branding artwork. |
| [compare-directories.png](compare-directories.png) | Static dialog/control illustration; review age before using it as current appearance evidence. |
| [file-properties.png](file-properties.png) | Static dialog/control illustration; review age before using it as current appearance evidence. |
| [monitor.png](monitor.png) | Static dialog/control illustration; review age before using it as current appearance evidence. |
| [preview-pane-image.png](preview-pane-image.png) | Generated sample content in a viewer/preview. |
| [theme-button-states-after-fix.png](theme-button-states-after-fix.png) | Native control contrast/state audit. |
| [theme-controls-catppuccin-frappe.png](theme-controls-catppuccin-frappe.png) | Native control/theme gallery; no identifying input fields. |
| [theme-controls-catppuccin-latte.png](theme-controls-catppuccin-latte.png) | Native control/theme gallery; no identifying input fields. |
| [theme-controls-catppuccin-mocha.png](theme-controls-catppuccin-mocha.png) | Native control/theme gallery; no identifying input fields. |
| [theme-controls-dark.png](theme-controls-dark.png) | Native control/theme gallery; no identifying input fields. |
| [theme-controls-dracula.png](theme-controls-dracula.png) | Native control/theme gallery; no identifying input fields. |
| [theme-controls-forest-mist.png](theme-controls-forest-mist.png) | Native control/theme gallery; no identifying input fields. |
| [theme-controls-high-contrast.png](theme-controls-high-contrast.png) | Native control/theme gallery; no identifying input fields. |
| [theme-controls-light.png](theme-controls-light.png) | Native control/theme gallery; no identifying input fields. |
| [theme-controls-neon-tokyo.png](theme-controls-neon-tokyo.png) | Native control/theme gallery; no identifying input fields. |
| [theme-controls-paper-and-ink.png](theme-controls-paper-and-ink.png) | Native control/theme gallery; no identifying input fields. |
| [theme-controls-rainbow-dark.png](theme-controls-rainbow-dark.png) | Native control/theme gallery; no identifying input fields. |
| [theme-controls-rainbow-light.png](theme-controls-rainbow-light.png) | Native control/theme gallery; no identifying input fields. |
| [theme-controls-retro-terminal.png](theme-controls-retro-terminal.png) | Native control/theme gallery; no identifying input fields. |
| [theme-controls-solar-flare.png](theme-controls-solar-flare.png) | Native control/theme gallery; no identifying input fields. |
| [theme-controls-ugly.png](theme-controls-ugly.png) | Native control/theme gallery; no identifying input fields. |
| [viewer-sqlite.png](viewer-sqlite.png) | Generated sample content in a viewer/preview. |

## Retired images

The October 19 review removes 43 image assets from `docs/res` and three obsolete
File Operations illustrations from `Specs/UI/Images`. Reasons include actual
workspace/test-run paths, real filenames, visible desktop/taskbar content, obsolete
pre-calculation/Mouse/editor UI, clipped legacy dialogs and an unused one-pixel placeholder. Current guide
embeds are removed; no redaction or substitute placeholder is published. Originals
and their metadata are preserved in a hash-verified private maintenance archive.
Do not restore retired assets merely to fill an illustration gap.

Three older safe configuration duplicates are also retired after preservation because
the current gallery now replaces them. Some previously retired filenames are reused
only for fresh native captures; original bytes remain in the private archive.

The Image/RAW legacy sample was also replaced with its already archived native
generated-bitmap capture after the provenance guard detected a byte mismatch.
Its original native scenario and provenance remain unchanged.

## Capture backlog

Written coverage exists for these workflows. Remaining illustrations need private,
deterministic scenarios and complete-frame provenance before publication:

- Full main-window chrome, navigation/history/drive menus, thumbnails, sorting,
  filtering, quick search and embedded text preview. Current workspace frames cover
  Brief, Detailed and Extra Detailed native panes.
- Further configured shortcut scopes, Terminal history suggestions and live-shell interaction.
- ShellNew and viewer menus beyond the captured canceled utility forms.
- Further protocol-specific remote workflows and playback controls/content varieties.
  The current provider frames use owned fakes; live accounts/devices are outside
  synthetic screenshot coverage. The current playback/pause frames use a real decoder.
- Text Diff, further SQLite states and additional image/RAW/PE content varieties.

Render viewer examples in the actual viewer rather than substituting Preferences.
Do not read a real account, device, folder history or file collection for a capture.
Static capture defaults to no activation. Interaction captures require an agreed
time and the existing runner warning/lease with foreground/cursor restoration.

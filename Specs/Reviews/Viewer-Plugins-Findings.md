> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


# Viewer plugins - verified findings (compact)

- **[security/CONFIRMED]** `Plugins/ViewerVLC/ViewerVLC.cpp:5014` - Process-global SetDllDirectoryW held for entire playback and raced by concurrent async loads, corrupting the DLL search path
- **[security/CONFIRMED]** `Plugins/ViewerWeb/ViewerWeb.cpp:2709` - Web-kind viewer loads untrusted local HTML with file:// origin, full script, and no CSP — only top-level http(s) is gated
- **[security/CONFIRMED]** `Plugins/ViewerWeb/ViewerWeb.cpp:3197` - NavigationStarting policy does not cover iframe/subframe navigations or new-window requests (no FrameNavigationStarting / NewWindowRequested handlers)
- **[security/CONFIRMED]** `Plugins/ViewerWeb/ViewerWeb.h:89` - allowExternalNavigation defaults to true, so a viewed file can navigate the pane to attacker-controlled remote URLs out of the box
- **[security/CONFIRMED]** `Plugins/ViewerWeb/ViewerWeb.cpp:3234` - JSON/Markdown external links bypass the allowExternalNavigation setting via unconditional ShellExecute
- **[security/CONFIRMED]** `Plugins/ViewerWeb/ViewerWeb.cpp:3219` - Web kind allows unconditional file:// (incl. UNC) navigation driven by viewed HTML content
- **[crash/PLAUSIBLE]** `RedSalamander/ViewerPluginManager.cpp:678` - directory_iterator throwing operator++ inside noexcept Discover() terminates the process
- **[dos-oom/CONFIRMED]** `Plugins/ViewerText/ViewerText.Text.cpp:2388` - Unbounded visual-line expansion on a single very long line (OOM / multi-second UI hang)
- **[dos-oom/CONFIRMED]** `Plugins/ViewerText/ViewerText.Text.cpp:3477` - Synchronous full-chunk decode + line/visual indexing on the UI thread (ANR on large chunks)
- **[dos-oom/CONFIRMED]** `Plugins/ViewerSpace/ViewerSpace.cpp:7012` - Unbounded _nodes vector resize keyed on attacker-influenced node id runs under a noexcept drain (bad_alloc -> std::terminate / OOM)
- **[dos-oom/CONFIRMED]** `Plugins/ViewerSpace/ViewerSpace.cpp:6478` - Scan node tree has no depth/node cap; cyclic or hostile IFileSystem plugin drives unbounded node growth to OOM
- **[dos-oom/CONFIRMED]** `Plugins/ViewerImgRaw/ViewerImgRaw.Decode.cpp:742` - Non-RAW WIC decode path has no megapixel/decompression-bomb bound below 16384x16384 -> ~1 GiB allocation
- **[dos-oom/CONFIRMED]** `Plugins/ViewerVLC/ViewerVLC.cpp:5121` - Synchronous libvlc.dll load and VLC install auto-detection block the UI thread
- **[dos-oom/CONFIRMED]** `Plugins/ViewerPE/ViewerPE.cpp:1833` - PE parse worker is join()ed on the UI thread during Close/navigate/destroy, freezing the UI until a possibly-hung remote Read returns
- **[dos-oom/CONFIRMED]** `Plugins/ViewerWeb/ViewerWeb.cpp:4167` - Web viewer copies entire virtual-FS file to a temp file with no size cap (disk-fill DoS)
- **[dos-oom/PLAUSIBLE]** `Plugins/ViewerWeb/ViewerWeb.cpp:4026` - UTF-16 BOM decode path materializes a full second wide-char copy with no size guard
- **[data-loss/CONFIRMED]** `Plugins/ViewerText/ViewerText.cpp:8665` - Save As with re-encoding silently writes only the loaded stream chunk for large (streamed) files
- **[leak/CONFIRMED]** `Plugins/ViewerSqlite/ViewerSqlite.Engine.cpp:432` - Temp DB snapshot created without delete-on-close leaks plaintext copies of sensitive databases on crash
- **[leak/PLAUSIBLE]** `Plugins/ViewerWeb/ViewerWeb.cpp:2695` - Extracted temp file leaks when WebView2 still holds it open across rapid file switches / abnormal exit
- **[incomplete-rendering/CONFIRMED]** `Plugins/ViewerText/ViewerText.cpp:5235` - DBCS/multibyte character split across stream-chunk boundary is corrupted (no carry for non-UTF8 code pages)
- **[incomplete-rendering/CONFIRMED]** `Plugins/ViewerText/ViewerText.Hex.cpp:2608` - Astral (non-BMP) code points in hex-view UTF-8 decode are replaced with U+FFFD instead of a surrogate pair
- **[robustness/CONFIRMED]** `Plugins/ViewerText/ViewerText.cpp:5582` - Hex-fallback-on-text-load-failure block is dead code (always-true S_OK precedes the FAILED() guard)
- **[robustness/CONFIRMED]** `Plugins/ViewerSqlite/ViewerSqlite.Engine.cpp:182` - Direct read-only open of the user's original DB locks/mutates it; valid WAL-mode databases fail to open
- **[robustness/CONFIRMED]** `Plugins/ViewerSpace/ViewerSpace.cpp:6457` - Misaligned NextEntryOffset lets crafted directory buffers reinterpret unaligned bytes as 64-bit FileInfo fields
- **[robustness/CONFIRMED]** `Plugins/ViewerSpace/ViewerSpace.cpp:6474` - Reparse-point directories are skipped entirely, so junctions/symlinks/mount points silently drop their target subtree's size from the total
- **[robustness/CONFIRMED]** `Plugins/ViewerSpace/ViewerSpace.cpp:6498` - EndOfFile sign check accepts negative size from plugin and silently treats as zero, but the >0 guard relies on signed comparison of file-controlled value
- **[robustness/CONFIRMED]** `Plugins/ViewerImgRaw/ViewerImgRaw.Decode.cpp:149` - EXIF IFD entry value type/count ignored: SHORT read from a tag declared as 1 short but with attacker count causes only first element read; tag type confusion can mis-parse ISO/orientation but stays in-bounds (defensive note, not OOB)
- **[robustness/CONFIRMED]** `Plugins/ViewerImgRaw/ViewerImgRaw.Export.cpp:397` - Temp file is created in the destination directory, which can fail or leak when that directory differs in writability/length from a long destination path
- **[robustness/CONFIRMED]** `RedSalamander/ViewerPluginManager.cpp:735` - Conflict-unload of one logical plugin unregisters localization for still-loaded sibling plugins from the same multi-plugin DLL
- **[robustness/PLAUSIBLE]** `Plugins/ViewerSpace/ViewerSpace.cpp:6478` - Real directory node IDs are unbounded and collide with the file-record / synthetic ID namespaces (0x40000000 / 0x80000000)
- **[robustness/PLAUSIBLE]** `Plugins/ViewerSpace/ViewerSpace.cpp:6519` - scannedFiles / scannedFolders / per-directory otherCount are uint32_t and overflow on huge trees, corrupting displayed and aggregate counts
- **[robustness/PLAUSIBLE]** `Plugins/ViewerSpace/ViewerSpace.cpp:6457` - Directory enumeration buffer parsed via unaligned reinterpret_cast of FileInfo at arbitrary offset
- **[robustness/PLAUSIBLE]** `Plugins/ViewerImgRaw/ViewerImgRaw.Decode.cpp:1116` - Embedded RAW thumbnail buffer (thumb.thumb / thumb.tlength) is trusted as the JPEG/bitmap length without independent validation against the source file bounds
- **[robustness/PLAUSIBLE]** `Plugins/ViewerVLC/ViewerVLC.cpp:4972` - extraArgs / videoOutput / audioOutput configuration is passed verbatim as libVLC instance arguments
- **[robustness/PLAUSIBLE]** `RedSalamander/ViewerPluginManager.cpp:801` - Unbounded trust of plugin-reported count drives OOB reads over metaData[i]
- **[contract/CONFIRMED]** `Plugins/ViewerText/ViewerText.cpp:5326` - Error log dereferences result->hr before it is assigned in MultiByteToWideChar failure path
- **[contract/PLAUSIBLE]** `Plugins/ViewerSqlite/ViewerSqlite.cpp:1036` - Sort-column index validated only against the displayed grid's column count, then emitted as a raw ORDER BY ordinal

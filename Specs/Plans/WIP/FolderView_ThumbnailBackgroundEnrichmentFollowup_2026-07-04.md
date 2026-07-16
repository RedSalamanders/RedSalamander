# FolderView Thumbnail Background Enrichment Follow-up (2026-07-04)

## Decision

Ship the cached-only visible thumbnail path. The visible FolderView thumbnail path remains fast-first:

- local files request shell thumbnails with `SIIGBF_INCACHEONLY`;
- likely WIC-supported image files may decode a bounded first-frame thumbnail through WIC after a shell-cache miss;
- all other cold-cache misses render the normal icon in the thumbnail slot;
- provider-allowed shell thumbnail extraction MUST NOT run on the visible path.

This intentionally accepts that Release users can initially see generic icons for cold-cache non-WIC formats such as video, PDF, and Office documents where the older synchronous/provider-allowed path could produce a real preview. That loss of completeness is accepted only because background enrichment is now a real follow-up, not a vague cleanup note.

## Scope

Implement a background enrichment path that can improve cold-cache non-WIC thumbnails after the first fast paint without blocking navigation, scrolling, close, resize, display-mode changes, or pane refresh.

In scope:

- Local-shell-backed file systems only.
- Items currently visible or near-visible in thumbnail mode.
- Cold-cache non-WIC misses that rendered fallback icons.
- Provider-allowed shell thumbnail extraction on bounded background workers.
- Late result application through the existing posted-payload path with stale generation checks.
- TW-2/TW-3/TW-4 cleanup: make the provider-allowed machinery reachable for this background path, fix the deadline ownership race, and make pending bitmap accounting exact.

Out of scope:

- Native shell/WIC thumbnail extraction for virtual, archive, remote-plugin, cloud-plugin, MTP/PTP, or otherwise non-local provider paths.
- Synchronous provider-allowed shell extraction during visible item layout/paint.
- Waiting for enrichment during pane close, navigation away, directory refresh, display-mode switch, or thumbnail-size change.
- A new durable app-owned thumbnail database unless separately planned and perf-validated.

## User Contract

- First paint and first visible thumbnail population prioritize responsiveness over completeness.
- Fallback icons are acceptable temporarily for cold-cache non-WIC formats.
- If background enrichment later produces a thumbnail for the same item generation and size, the pane may replace the fallback icon in place and invalidate only the affected item/area.
- Stale enrichment results are silent no-ops. A thumbnail generated for an old folder, generation, sort, item index, or thumbnail size MUST NOT update the current view.
- Users should not need to press Refresh to benefit from enrichment for currently visible items.

## Implementation Requirements

- Keep the normal visible path cached-only and keep `folderView_thumbnail_cached_only_no_close_stall` green.
- Enqueue enrichment only after a visible thumbnail request has completed with an icon fallback because cached shell lookup missed and WIC fallback did not apply or failed.
- Capture only values in worker payloads: `HWND`, pane/view generation, thumbnail batch/generation, item index or stable item key, normalized local path, target size, and request reason. Do not capture `this`.
- Initialize COM on enrichment workers as MTA.
- Bound work with both a per-request deadline and queue/backpressure caps. A slow shell provider must age out as timeout/cancel, not hold a worker indefinitely.
- Treat worker completion as abandonable. Closing or navigating the pane must only mark generations stale/cancel future queueing; it must not join on enrichment workers.
- Post late results through the standard posted-payload mechanism. The UI-thread receiver must re-check `HWND`, generation, item identity, thumbnail size, display mode, and current local-shell-backed eligibility before applying.
- Every posted bitmap-creation payload must have matching pending-counter accounting, including abandoned/late-delivery paths and post-failure paths.
- The deadline loop must handle the boundary case where extraction completes at the deadline: do one final zero-timeout completion poll before declaring timeout, or otherwise make callback/waiter ownership unambiguous.
- Provider-allowed extraction failures and timeouts are normal control flow. Emit metrics, but do not surface pane errors or log noisy warnings for ordinary misses.

## Metrics And Diagnostics

The follow-up must preserve existing thumbnail diagnostics and add enough evidence to separate visible fast path from background enrichment:

- `thumbnails.cached_extract_us`
- `thumbnails.shell_cache_hit_count`
- `thumbnails.shell_cache_miss_count`
- `thumbnails.shell_provider_allowed_count`
- `thumbnails.shell_provider_timeout_count`
- `thumbnails.close_to_idle_us`
- add or reuse counters for enrichment queued, completed, stale-dropped, timed-out, applied, and canceled.

Metric rows must distinguish:

- visible cached-only request,
- WIC fallback,
- background provider-allowed enrichment,
- late stale drop,
- timeout.

## Required Tests

- Fast-first guard: cold-cache non-WIC fixture still renders fallback without running provider-allowed work on the visible path.
- Enrichment success: a cold-cache non-WIC fixture queues background provider-allowed extraction, applies the late thumbnail for the same generation, and repaints only the affected item/area.
- Stale navigation: navigate away or refresh before enrichment completes; late result is dropped and pending counters return to zero.
- Close/no-join: force a slow provider-allowed extraction and close the pane/window; close latency stays below the existing bounded deadline and no payload leaks.
- Deadline boundary: inject completion at the deadline edge and prove the result is either delivered exactly once or timed out without orphaning ownership.
- Pending accounting: every synchronous, timeout, abandoned, late-posted, stale-dropped, and post-failure path returns pending bitmap count to zero.
- Non-local guard: virtual/remote/archive/cloud/MTP paths never call shell/WIC thumbnail extraction and keep icon fallback.

## Performance Gate

Because FolderView is a hot path, implementation must archive before/after evidence under `Specs/TestRuns/`:

- thumbnail-mode cold open with mixed local non-WIC files,
- scroll while enrichment is active,
- navigation/close while a slow provider is injected,
- warm-cache comparison after enrichment has populated the shell cache.

Passing evidence must show:

- first visible response remains governed by cached-only path,
- provider-allowed enrichment does not increase close/navigation latency beyond the existing gate,
- queue caps prevent unbounded worker growth,
- stale-drop counts are visible when churn is induced,
- no sustained UI-thread frame regression from applying late thumbnails.

## Closeout

When implemented:

- update `Specs/UI/UI_FolderView.md` with the final enrichment behavior and metrics;
- update or retire the TW-2/TW-3/TW-4 rows in the Tailwind plan;
- keep the Granite GR-D1 decision row marked as signed off;
- archive test/perf evidence under `Specs/TestRuns/`;
- move this plan to `Specs/Plans/Done/`.

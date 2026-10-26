# IMAP Listing Metadata Repair Evidence

Date: 2026-05-14
Branch: `master`
Configuration: Debug x64

## Scenario

This run validates the deterministic portion of the pane-listing metadata repair:

- directory listings still prefer bulk `UID FETCH` summary requests;
- if a bulk request fails or misses requested UIDs, missing summaries are retried in bounded 16-UID batches;
- if a repair batch still misses a UID, the listing path falls back to a single-UID summary fetch before showing numeric `<uid>.eml`, `0 B`, or missing timestamps.

The screenshot-driven live issue was that Properties could fetch metadata for one UID while the pane listing showed numeric names and `0 B`. No deterministic live IMAP fixture credentials are available in this workspace, so the guard is a focused repair-plan test plus the existing production diagnostics.

## Red Check

```text
.\build.ps1 -ProjectName FileSystemCurlTests
```

Observed before implementation:

```text
FileSystemCurlTests.cpp(...): error C2039: 'ImapUidBatchRange': is not a member of 'FileSystemCurlInternal'
FileSystemCurlTests.cpp(...): error C2039: 'BuildImapUidBatchRanges': is not a member of 'FileSystemCurlInternal'
```

## Commands

```powershell
.\build.ps1 -ProjectName FileSystemCurlTests
.\.build\x64\Debug\FileSystemCurlTests.exe
.\.build\x64\Debug\FileSystemCurlTests.exe --perf
```

## Results

```text
FileSystemCurlTests: 0 warning(s), 0 error(s)
FileSystemCurlTests passed.

parser iterations=500000 elapsedUs=89966 checksum=61728442388895
builder iterations=500000 elapsedUs=1932667
statusParser iterations=500000 elapsedUs=712806
summaryRepairBatches missing=37 maxBatch=16 batches=3
messagePropertiesCommands messages=1 baseline=5 candidate=4 reductionPercent=20.0
messagePropertiesCommands messages=200 baseline=5 candidate=4 reductionPercent=20.0
messagePropertiesCommands messages=201 baseline=6 candidate=4 reductionPercent=33.3
messagePropertiesCommands messages=1000 baseline=9 candidate=4 reductionPercent=55.6
messagePropertiesCommands messages=10000 baseline=54 candidate=4 reductionPercent=92.6
```

## Metrics

The production listing path now emits `filesystem.imap.summary_missing_count` alongside the existing:

- `filesystem.imap.fetch_summaries_us`
- `filesystem.imap.summary_response_bytes`
- `filesystem.imap.summary_parse_us`
- `filesystem.imap.summary_repair_count`
- `filesystem.imap.build_fileinfo_us`

# IMAP Metadata Display And Properties Performance Closeout

Status: Done for the deterministic code, test, spec, and archived-evidence scope on 2026-05-14.

## User Contract Delivered

- IMAP message names now use `<subject> [<uid>].eml`.
- The UID remains the operation key and is parsed from the final bracketed suffix.
- Direct numeric paths such as `12345.eml` keep working.
- Ambiguous subject-plus-trailing-digits names such as `Quarterly report 12345.eml` are rejected.
- Long subjects are sanitized and capped at 96 UTF-16 code units with ASCII `...` before the UID suffix.
- IMAP subjects are decoded from RFC2047 Q/B encoded words across common MIME charset aliases before filename sanitizing; malformed `=_charset_q_...` fragments are recovered and soft hyphens are made visible as `-`.
- The existing IMAP listing mapping continues to expose received/internal date as `LastWriteTime`/`ChangeTime` and RFC822 size as file size.
- IMAP pane listings repair missing bulk summary metadata with bounded smaller UID batches and single UID retries before falling back to numeric names or empty file metadata.
- IMAP message Properties avoid a full mailbox summary fetch for one selected message.
- IMAP mailbox folder Properties now expose mailbox path, server mailbox path, hierarchy delimiter, and `STATUS` counts when available.

## Code Changes

- `Plugins/FileSystemCurl/FileSystemCurl.ImapHelpers.h`
- `Plugins/FileSystemCurl/FileSystemCurl.ImapHelpers.cpp`
- `Plugins/FileSystemCurl/FileSystemCurl.Imap.cpp`
- `Plugins/FileSystemCurl/FileSystemCurl.Internal.h`
- `Plugins/FileSystemCurl/FileSystemCurl.vcxproj`
- `Tests/FileSystemCurlTests/FileSystemCurlTests.cpp`
- `Tests/FileSystemCurlTests/FileSystemCurlTests.vcxproj`
- `RedSalamander.sln`

## Tests

Added `FileSystemCurlTests`, a deterministic standalone test harness covering:

- preferred IMAP leaf-name formatting,
- UID parsing for preferred and direct names,
- RFC2047 subject decoding across UTF-8, base64, non-UTF code pages, mixed fragments, and malformed sanitized fragments,
- malformed name rejection,
- mailbox `STATUS` parsing,
- large missing-summary repair batch planning,
- message Properties command-count model.

Validation commands:

```powershell
.\build.ps1 -ProjectName FileSystemCurlTests
.\.build\x64\Debug\FileSystemCurlTests.exe
.\.build\x64\Debug\FileSystemCurlTests.exe --perf
.\build.ps1 -ProjectName FileSystemCurl
```

Observed result: all commands passed with 0 warnings and 0 errors.

## Performance Evidence

Archived under:

```text
Specs/TestRuns/4cb089111a23/FileSystemCurl/2026-05-14_imap_metadata_display_perf/
Specs/TestRuns/4cb089111a23/FileSystemCurl/2026-05-14_imap_listing_metadata_repair/
Specs/TestRuns/4cb089111a23/FileSystemCurl/2026-05-14_imap_subject_decoding/
```

Deterministic perf probe:

```text
parser iterations=500000 elapsedUs=90667 checksum=61728445488895
builder iterations=500000 elapsedUs=1870691
statusParser iterations=500000 elapsedUs=733122
subjectDecoder iterations=100000 elapsedUs=544695
summaryRepairBatches missing=37 maxBatch=16 batches=3
messagePropertiesCommands messages=1 baseline=5 candidate=4 reductionPercent=20.0
messagePropertiesCommands messages=200 baseline=5 candidate=4 reductionPercent=20.0
messagePropertiesCommands messages=201 baseline=6 candidate=4 reductionPercent=33.3
messagePropertiesCommands messages=1000 baseline=9 candidate=4 reductionPercent=55.6
messagePropertiesCommands messages=10000 baseline=54 candidate=4 reductionPercent=92.6
```

The proven performance win is for single-message Properties: the candidate path is constant with mailbox size and avoids the previous full-mailbox summary fetch before showing one message.

The pane-listing metadata repair guard proves that large missing summary sets are split into bounded repair batches instead of being skipped after the initial bulk summary request.

## Production Metrics Added

- `filesystem.imap.read_directory_us`
- `filesystem.imap.list_mailboxes_us`
- `filesystem.imap.list_uids_us`
- `filesystem.imap.fetch_summaries_us`
- `filesystem.imap.summary_response_bytes`
- `filesystem.imap.summary_parse_us`
- `filesystem.imap.summary_missing_count`
- `filesystem.imap.summary_repair_count`
- `filesystem.imap.build_fileinfo_us`
- `filesystem.imap.properties_message_us`
- `filesystem.imap.properties_mailbox_us`
- `filesystem.imap.status_mailbox_us`

## Durable Specs Updated

- `Specs/FileSystem/FileSystem_Imap.md`
- `Tests/README.md`

## Deferred Live Validation

Live IMAP wall-clock listing comparison is blocked in this workspace by missing deterministic IMAP fixture credentials. The code now emits the required `filesystem.imap.*` metrics so a future live run can compare listing latency and response bytes before changing the listing `FETCH` request shape.

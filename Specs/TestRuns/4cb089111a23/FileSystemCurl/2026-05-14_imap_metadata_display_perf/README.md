# IMAP Metadata Display And Properties Performance Evidence

Date: 2026-05-14
Branch: `codex/imap-metadata-display-perf`
Configuration: Debug x64

## Scenario

This run validates the deterministic, credential-free portion of the IMAP metadata display work:

- preferred message leaf names use `<subject> [<uid>].eml`;
- UID parsing accepts preferred and direct numeric names;
- malformed non-key names are rejected;
- mailbox `STATUS` responses parse without a live server;
- single-message Properties uses a constant-size command model rather than the previous full-mailbox summary path.

Live IMAP server latency was not used for this closeout because no deterministic IMAP fixture or credentials are available in this workspace. The production code now emits `filesystem.imap.*` perf metrics for live runs.

## Commands

```powershell
.\build.ps1 -ProjectName FileSystemCurlTests
.\.build\x64\Debug\FileSystemCurlTests.exe
.\.build\x64\Debug\FileSystemCurlTests.exe --perf
.\build.ps1 -ProjectName FileSystemCurl
```

## Results

```text
FileSystemCurlTests passed.

parser iterations=500000 elapsedUs=111134 checksum=61728442388895
builder iterations=500000 elapsedUs=2016913
statusParser iterations=500000 elapsedUs=734219
messagePropertiesCommands messages=1 baseline=5 candidate=4 reductionPercent=20.0
messagePropertiesCommands messages=200 baseline=5 candidate=4 reductionPercent=20.0
messagePropertiesCommands messages=201 baseline=6 candidate=4 reductionPercent=33.3
messagePropertiesCommands messages=1000 baseline=9 candidate=4 reductionPercent=55.6
messagePropertiesCommands messages=10000 baseline=54 candidate=4 reductionPercent=92.6
```

Build verification:

```text
FileSystemCurlTests: 0 warning(s), 0 error(s)
FileSystemCurl: 0 warning(s), 0 error(s)
```

## Command-Count Model

Baseline message Properties path:

```text
GetEntryInfo -> ReadDirectoryEntries(parent)
  LIST mailboxes
  UID SEARCH ALL
  UID FETCH summary chunks for every message in mailbox
Properties IMAP section
  LIST hierarchy delimiter
  UID FETCH selected message
```

Candidate message Properties path:

```text
GetEntryInfo
  LIST hierarchy delimiter
  UID FETCH selected message
Properties IMAP section
  LIST hierarchy delimiter
  UID FETCH selected message
```

The candidate is constant with respect to mailbox size. The previous path grew with `ceil(messageCount / 200)` summary fetches before it could show one message's Properties.

## Caveat

Directory-listing latency still requires a live or scripted IMAP fixture for true before/after wall-clock comparison. The code now emits:

- `filesystem.imap.read_directory_us`
- `filesystem.imap.list_mailboxes_us`
- `filesystem.imap.list_uids_us`
- `filesystem.imap.fetch_summaries_us`
- `filesystem.imap.summary_response_bytes`
- `filesystem.imap.summary_parse_us`
- `filesystem.imap.summary_repair_count`
- `filesystem.imap.build_fileinfo_us`
- `filesystem.imap.properties_message_us`
- `filesystem.imap.properties_mailbox_us`
- `filesystem.imap.status_mailbox_us`

Those metrics should be captured against a real mailbox before changing the listing `FETCH` request shape.

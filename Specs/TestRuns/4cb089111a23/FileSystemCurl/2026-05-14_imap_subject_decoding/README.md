# IMAP Subject Decoding Evidence

Date: 2026-05-14
Branch: `master`
Configuration: Debug x64

## Scenario

This run validates IMAP subject/title transcoding before pane filename construction and Properties display:

- standard RFC2047 encoded-words using Q and B encodings;
- mixed encoded and plain subject fragments;
- UTF-8 emoji and Japanese text;
- non-UTF Windows code page text;
- malformed sanitized fragments such as `=_utf-8_q_...`;
- U+00AD soft hyphen normalization to visible `-`.

## Red Check

```text
.\build.ps1 -ProjectName FileSystemCurlTests
```

Observed before implementation:

```text
error C2039: 'DecodeRfc2047EncodedWordsToUtf16': is not a member of 'FileSystemCurlInternal'
```

After the first implementation pass, the regression suite also exposed the previous parser bug where a Q-encoded word beginning with `=F0` was kept literal because the parser searched for the first `?=` before identifying the encoded-text boundary.

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

## Caveat

This is deterministic title-decoder coverage. Live mailbox validation still depends on fixture credentials, but the production pane and Properties paths now share the tested decoder before filename sanitizing.

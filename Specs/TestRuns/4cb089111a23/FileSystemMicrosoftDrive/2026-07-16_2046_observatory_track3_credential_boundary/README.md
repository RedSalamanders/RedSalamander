# Observatory Track 3 Microsoft Drive Credential-Boundary Evidence

Date: 2026-07-16
Machine profile: `4cb089111a23`
CPU: AMD Ryzen 9 9950X3D 16-Core Processor
OS: Windows 11 Pro 10.0.26200
Working-tree base: `f4e0c8c3bed8`
Configurations: Debug x64 deterministic selftests; Release x64 production build

## Scenario

Protect Graph bearer tokens, opaque paging continuations, and preauthenticated upload-session URLs at the
Microsoft Drive transport boundary. A valid same-origin `@odata.nextLink` must preserve normal two-page listing
behavior. Plaintext, foreign-origin, malformed, fragment-bearing, userinfo-bearing, and repeated continuations must
fail before an unauthorized extra request. Upload-session URLs must be HTTPS on the approved Microsoft service
origins, must never receive the Graph bearer token, and must not be followed through automatic redirects.

The user-visible performance contract is a request-count bound rather than a latency claim:

- approved two-page listing: exactly two Graph requests;
- foreign or plaintext continuation: exactly one Graph request, then local rejection;
- repeated second-page continuation: exactly two Graph requests, then local rejection before a third;
- credential-bound validation adds no network round trip and remains local to request dispatch.

## Commands and results

```powershell
.\build.ps1 -ProjectName FileSystemMicrosoftDrive -Configuration Debug
.\.build\x64\Debug\PluginContractTests.exe
Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
.\build.ps1 -ProjectName FileSystemMicrosoftDrive -Configuration Release
```

- Debug plugin build: 0 warnings / 0 errors.
- Release plugin build: 0 warnings / 0 errors.
- PluginContractTests: passed; Microsoft Drive debug selftests passed 103 / failed 0.
- TestHarnessSourceContracts: passed 133 / failed 0 / skipped 0.
- Injected Graph and upload transport failures emitted only redacted target classes; bearer and opaque-query
  sentinels were absent from the captured diagnostics.

## Performance interpretation

No performance improvement is claimed. The security boundary performs one local WinHTTP URL parse and bounded
origin/path checks before credential-bearing dispatch. Valid pagination preserves the expected request count;
rejected continuations avoid the unsafe network request entirely. Network latency was not sampled because this
workspace has no deterministic Microsoft Graph fixture or test credentials, and live service timing would not be a
stable regression gate.

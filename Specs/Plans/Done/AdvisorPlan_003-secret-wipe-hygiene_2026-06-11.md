# Advisor Plan 003 - Hoist SecureClear and wipe secrets consistently

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/003-secret-wipe-hygiene.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/003-secret-wipe-hygiene.md`
- **Live owner:** none (implemented; durable helper owned by Specs/Core/Core_SharedHelpers.md)
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 003: Hoist SecureClear into Common and wipe secrets consistently

> **Frozen historical instructions (do not execute)**: This plan is archived; do not follow these steps. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. Historical: status-row updates no longer apply for this plan
> in `plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat 2c41eb0b6..HEAD -- RedSalamander/ConnectionSecrets.cpp RedSalamander/ConnectionCredentialPromptDialog.cpp Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp Plugins/FileSystemS3 Common/Common`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

> **DONE — merged to master `2572c5681` ("Merge secret wipe hygiene"), 2026-06-20.** The advisor's verdict had been deferred at execution time because the full self-test suite couldn't run (machine-wide `TryAcquireSelfTestRunMutex` held by a concurrent agent). The maintainer merged the work; reconcile (2026-06-21, HEAD `a72512919`) re-verified every static done-criterion at HEAD and all pass: `SecureWipe::SecureClear` is the single shared helper in `Common/Helpers.h:466,476`; the MicrosoftDrive file-local copy is gone; FileSystemS3 has a `SecureClearSensitiveFields()` destructor-helper + 11 `SecureClear` uses; `ConnectionCredentialPromptDialog` wipes `_acceptedSecret` in its destructor with correct copy-out-then-wipe timing; and the three `g_quickConnect*` secret strings no longer use plain `.clear()`. The full-suite gate was the maintainer's merge-time call and was not re-run during reconcile.

- **Priority**: P1
- **Effort**: M
- **Risk**: LOW (additive wiping; no control-flow changes)
- **Depends on**: none
- **Category**: security
- **Planned at**: commit `2c41eb0b6`, 2026-06-12 (refresh of the original `d6bfccc42` plan; all excerpts and line anchors re-verified at this SHA — zero drift on the in-scope files in between)

## Why this matters

The codebase already has a correct best-effort secret-wiping pattern — `SecureClear` in the Microsoft Drive plugin (16 call sites) — but it is file-local, and the other secret holders don't use anything like it: quick-connect passwords/passphrases/refresh tokens live in global `std::wstring`s cleared with plain `.clear()`, the credential prompt dialog keeps the accepted secret in a plain member, and the S3 plugin never wipes the secret access key it fetches. Plain `.clear()` leaves the secret bytes in freed/retained heap memory, recoverable from process dumps. This plan makes the existing pattern shared and applies it consistently. Honest framing: `std::wstring` reallocation can still leave stray copies — this is defense-in-depth parity with the repo's own best practice, not a guarantee; do not oversell it in comments.

## Current state

- `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp:652-667` — the exemplar, currently file-local:
  ```cpp
  void SecureClear(std::string& text) noexcept
  {
      if (! text.empty())
      {
          SecureZeroMemory(text.data(), text.size());
          text.clear();
      }
  }

  void SecureClear(std::wstring& text) noexcept
  {
      if (! text.empty())
      {
          SecureZeroMemory(text.data(), text.size() * sizeof(wchar_t));
          text.clear();
      }
  }
  ```
- `RedSalamander/ConnectionSecrets.cpp:177-184` — the secret holders:
  ```cpp
  std::mutex g_quickConnectMutex;
  std::optional<Common::Settings::ConnectionProfile> g_quickConnectProfile;
  std::wstring g_quickConnectPassword;
  std::wstring g_quickConnectPassphrase;
  std::wstring g_quickConnectRefreshToken;
  ```
- `RedSalamander/ConnectionSecrets.cpp:336-370` — `SetQuickConnectSecret` clears with `.clear()` (lines 345/349/353) and overwrites with `.assign(secret)` (lines 362/366/369+). Both paths leave old bytes behind: `.clear()` doesn't zero, and `.assign()` may reallocate, abandoning the old buffer unzeroed. There is also a `TryAssign` helper at lines 189-192 (plain `assign`).
- `RedSalamander/ConnectionCredentialPromptDialog.cpp:244-245` — dialog class members (class is defined in the .cpp):
  ```cpp
  std::wstring _acceptedUserName;
  std::wstring _acceptedSecret;
  ```
  `_acceptedSecret` holds the password the user typed; nothing wipes it when the dialog object is destroyed.
- `Plugins/FileSystemS3/FileSystemS3.Shared.cpp:478-510` — `out.secretAccessKey` (a `std::optional<std::string>` on the AWS context struct) is populated from the host secret store via `Utf8FromUtf16(secret.get())`; grep confirms **zero** `SecureClear`/`SecureZeroMemory` occurrences in the whole `Plugins/FileSystemS3` directory. Note the intermediate `wil::unique_cotaskmem_string secret` is freed but not zeroed either, and the `secretUtf8` local is a third copy.
- The context struct owning `accessKeyId`/`secretAccessKey` is declared in **`Plugins/FileSystemS3/FileSystemS3.Internal.h:71`** (verified). Credential consumption sites — `FileSystemS3.Shared.cpp:694-719` (AWS SDK config), `FileSystemS3.IO.cpp:1565`, `FileSystemS3.S3.cpp:138` — are **out of scope** (consumers, not owners; see Scope).
- Repo conventions that apply (AGENTS.md): WIL RAII, `noexcept` helpers, no `catch (...)`, clang-format via `format-all.ps1`. New shared helpers belong in `Common\Common\` (Helpers.h is the high-fan-in utility header).

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Build all | `.\build.ps1` | exit 0 |
| Build one project | `.\build.ps1 -ProjectName RedSalamander` | exit 0 |
| Selftests (canonical runner) | `.\Tools\Run-AllTests.ps1 -SkipBuild` | exit 0, summary shows all three suites passed (parses results.json + case coverage; documented in README "Self-tests") |
| One suite only (faster iteration) | `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2.0` | exit 0 |
| Format | `.\format-all.ps1` | exit 0 (requires clang-format; CI also auto-formats) |

## Scope

**In scope** (the only files you should modify):
- `Common\Common\Helpers.h` (or a new small header `Common\Common\SecureClear.h` included from the four consumers — your choice; prefer Helpers.h only if it already has similar small inline helpers, otherwise the new header keeps the junk-drawer from growing)
- `RedSalamander\ConnectionSecrets.cpp`
- `RedSalamander\ConnectionCredentialPromptDialog.cpp`
- `Plugins\FileSystemMicrosoftDrive\FileSystemMicrosoftDrive.cpp` (delete the local definitions, include the shared header)
- `Plugins\FileSystemS3\FileSystemS3.Shared.cpp` and `Plugins\FileSystemS3\FileSystemS3.Internal.h` (the context struct lives at `FileSystemS3.Internal.h:71`)

**Out of scope** (do NOT touch):
- Callers of `GetQuickConnectSecret` and every place secrets are *consumed* (curl options, AWS SDK config) — wiping consumer-side copies is a larger follow-up; this plan only fixes the owning stores.
- `WindowsHello.cpp`, Credential Manager interop (`CredDeleteW` etc. in ConnectionSecrets.cpp) — storage backend is fine; do not refactor it.
- Any behavior change: no API signatures change, no new prompts, no logging of secret-adjacent data.

## Git workflow

- Branch: `advisor/003-secret-wipe-hygiene`
- Commits: one for the shared helper + MicrosoftDrive switch, one for ConnectionSecrets/dialog, one for S3. Subject style: "Wipe quick-connect secrets with SecureClear" etc.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Create the shared helper

Add `SecureClear(std::string&)` and `SecureClear(std::wstring&)` as `inline ... noexcept` to the chosen Common header (exact bodies from the excerpt above — they are already correct). Namespace it consistently with neighboring helpers in that header (check what Helpers.h uses; if free functions at namespace scope are the norm, match that). Add one comment line: best-effort wipe; reallocation may leave copies.

**Verify**: `.\build.ps1 -ProjectName Common` → exit 0.

### Step 2: Switch FileSystemMicrosoftDrive to the shared helper

Delete the two local `SecureClear` definitions at `FileSystemMicrosoftDrive.cpp:652-667`, include the shared header, confirm all 16 existing call sites still compile unchanged.

**Verify**: `.\build.ps1 -ProjectName FileSystemMicrosoftDrive` → exit 0 (if per-project build doesn't accept plugin names, build the solution). `Grep "void SecureClear" Plugins\FileSystemMicrosoftDrive` → no matches.

### Step 3: ConnectionSecrets.cpp

In `SetQuickConnectSecret` (lines 336-370): replace each `.clear()` with `SecureClear(...)`, and in each assign branch call `SecureClear(target)` immediately before `.assign(secret)` so the old buffer is zeroed before a potential reallocation. Apply the same wipe-before-assign inside `TryAssign` (lines 189-192) **only if** `TryAssign` is used exclusively for secret fields — check its call sites first; if it's also used for non-secret strings, leave it alone and wipe at the three secret call sites instead. Also find where the quick-connect state is torn down/cleared wholesale (search the file for where `g_quickConnectProfile` is reset) and ensure the three secret strings get `SecureClear` there too.

**Verify**: `.\build.ps1 -ProjectName RedSalamander` → exit 0. `Grep -n "g_quickConnectPassword.clear|g_quickConnectPassphrase.clear|g_quickConnectRefreshToken.clear" RedSalamander\ConnectionSecrets.cpp` → no matches.

### Step 4: ConnectionCredentialPromptDialog.cpp

Wipe `_acceptedSecret` (and `_acceptedUserName` only where it is documented/used as secret-adjacent — if unsure, wipe just `_acceptedSecret`) in the dialog's destructor; if the class has no destructor, add one. Also wipe before any re-assignment of `_acceptedSecret` (find assignments: `Grep -n "_acceptedSecret" RedSalamander\ConnectionCredentialPromptDialog.cpp`). Mind the wipe-timing hazard: the accepted values are copied OUT to the caller at `ConnectionCredentialPromptDialog.cpp:407-408` (`userNameOut = _acceptedUserName; secretOut = _acceptedSecret;`) — the wipe must happen only in destruction/reset paths that run AFTER that copy-out, never before it.

**Verify**: build exits 0; the credential prompt still works — covered by the Commands selftest (connections cases): `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2.0` → exit 0.

### Step 5: FileSystemS3

The struct owning `accessKeyId`/`secretAccessKey` is at `Plugins\FileSystemS3\FileSystemS3.Internal.h:71` (cross-check with `Grep -rn "secretAccessKey" Plugins\FileSystemS3` — expect hits in Internal.h:71, Shared.cpp:509/694-719, IO.cpp:1565, S3.cpp:138; only the Internal.h declaration and the Shared.cpp assignment sites are yours to touch). Preferred fix: give that struct a destructor that `SecureClear`s both fields (and wipe-before-overwrite in `ResolveAwsContext` where they're assigned). Before adding a destructor, check how the struct is used: if it is copied around or stored long-term, adding a destructor is still safe (it remains copyable), but every copy multiplies secret instances — note that in your report rather than redesigning ownership here. Also wipe the intermediate `secretUtf8` local in `FileSystemS3.Shared.cpp:504-509` after it is moved/copied into `out.secretAccessKey` (assign via `std::move` if not already, then `SecureClear` the local if it still owns data — simplest: build into the destination directly, or wipe the local unconditionally after assignment).

**Verify**: `.\build.ps1` (full solution) → exit 0.

### Step 6: Format + full verification

`.\format-all.ps1`, rebuild, then run the full selftest suite: `.\Tools\Run-AllTests.ps1 -SkipBuild -TimeoutMultiplier 2.0` → exit 0, all three suites green.

## Test plan

There is no practical unit test for "memory was zeroed" without UB-adjacent inspection; the regression net is:
- Full selftest run (`--selftest` exit 0) — the commands suite exercises connection prompts and settings paths.
- Grep gates in Done criteria preventing the plain-`.clear()` pattern from coming back at these sites.
No new test files are required for this plan.

## Done criteria

- (archived, not live)  `SecureClear` defined exactly once in `Common\` (grep: `Grep -rn "void SecureClear" --glob "!*.md"` → 1 header, 0 plugin-local definitions)
- (archived, not live)  `Grep -n "\.clear\(\)" RedSalamander\ConnectionSecrets.cpp` shows no remaining `.clear()` on the three `g_quickConnect*` secret strings
- (archived, not live)  ConnectionCredentialPromptDialog wipes `_acceptedSecret` in its destructor
- (archived, not live)  `Grep -rn "SecureClear" Plugins\FileSystemS3` → at least 2 matches
- (archived, not live)  `.\build.ps1` exit 0; `.\Tools\Run-AllTests.ps1 -SkipBuild -TimeoutMultiplier 2.0` exit 0 (all three suites)
- (archived, not live)  No files outside the in-scope list modified (`git status`)
- (archived, not live)  `plans/README.md` status row updated

## STOP conditions

Stop and report back if:

- The S3 context struct turns out to be shared/cached across operations with unclear ownership (e.g. stored in a long-lived map) — wiping in a destructor could then zero a secret still in use elsewhere only if something aliases the buffer (it shouldn't with value semantics, but if you see pointer/reference aliasing of these fields, stop).
- `TryAssign` in ConnectionSecrets.cpp is used for non-secret strings AND refactoring around it would touch out-of-scope files.
- Any selftest that passed at baseline fails after your change.
- You find the secrets are *also* written to logs/trace anywhere you touch — do not fix silently; report it (separate finding).

## Maintenance notes

- New secret-holding members anywhere in the codebase should use this helper; consider an AGENTS.md "Regression Guards" bullet as a follow-up (deferred — docs guideline change is out of scope here).
- Reviewer focus: Step 4's wipe-timing (must not wipe before the caller consumes the accepted secret) and that Step 5 didn't change S3 context copy semantics.
- Deferred: wiping consumer-side copies (curl/AWS SDK internals hold their own copies regardless); a dedicated non-reallocating secure-string type.

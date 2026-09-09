# SelfTest Remote Credentials (Optional)

## Per-machine resource manifest

The canonical machine-specific input is the optional file
`X:\RedSalamander.Perf\config\machine-resources.json`, validated by
`Specs/Testing/TestMachineResources.schema.json`. `X:` is the same fixed local drive
selected as `REDSALAMANDER_TEST_ROOT`; the manifest path is fixed and cannot be
redirected outside that root.

The manifest stores only non-secret selectors:

- Connection Manager profile names for FTP, SFTP, SCP, IMAP, AWS-compatible S3,
  OneDrive Personal/Business, and SharePoint;
- one dedicated SMB UNC self-test directory;
- an approved MTP device/profile/root/scratch selector;
- one exact alternate fixed-local-drive root for cross-volume or filesystem-capability FileOperations coverage.

Passwords, OAuth refresh tokens, key passphrases, AWS access keys, and other secrets
MUST NOT appear in this JSON. Connection Manager continues to retrieve them from
Windows Credential Manager. When the manifest exists, it is authoritative: the runner
clears all managed legacy resource environment variables, projects only enabled
resources for the duration of the run, fingerprints a SHA-256 capability digest, and
restores the caller environment afterward. When the manifest is absent, the historical
environment variables below remain supported for compatibility.

Example:

```json
{
  "schema": "red-salamander.test-machine-resources.v1",
  "resources": [
    {
      "id": "connection.s3.primary",
      "kind": "connection-profile",
      "enabled": true,
      "profile": "SINON S3 SelfTest"
    },
    {
      "id": "connection.onedrive.personal",
      "kind": "connection-profile",
      "enabled": true,
      "profile": "SINON OneDrive SelfTest"
    },
    {
      "id": "path.smb.primary",
      "kind": "smb-share",
      "enabled": true,
      "root": "\\\\server\\share\\RedSalamander-SelfTest"
    },
    {
      "id": "path.local.alternate",
      "kind": "local-test-root",
      "enabled": true,
      "root": "D:\\RedSalamander.Perf"
    }
  ]
}
```

Resource IDs are stable test contracts. Adding a new provider requires extending the
schema and projection allow-list in the same change; arbitrary environment-variable
names are deliberately not accepted.

`path.local.alternate` selects one exact `<AltDrive>:\RedSalamander.Perf` root.
It does not authorize arbitrary directories or enable cross-volume mutation by itself:
the selected drive must be fixed and different from the primary run volume, and the
caller must still pass `Tools\Run-AllTests.ps1 -AllowAlternateVolumeTestRoot` for that
run. When configured, FileOperations probes only this drive instead of enumerating
other fixed volumes. All generated data remains below its marked
`runs\<runId>\scratch` subtree.

`RedSalamander.exe --fileops-selftest` includes **Phase 16** checks that validate the host can retrieve **saved secrets** (passwords / SSH key passphrases / OAuth refresh tokens) for remote filesystem plugins:

- FTP (`builtin/file-system-ftp`)
- SFTP (`builtin/file-system-sftp`)
- SCP (`builtin/file-system-scp`)
- IMAP (`builtin/file-system-imap`)
- S3 (`builtin/file-system-s3`)
- OneDrive Personal (`builtin/file-system-onedrive-personal`)
- OneDrive Business (`builtin/file-system-onedrive-business`)
- SharePoint (`builtin/file-system-sharepoint`)

These phases are **conditional**:

- If the connection profile and secret are present, the phase **passes**.
- Otherwise the phase is **skipped** (selftests stay green by default).

All such cases remain declared members of the suite and must still emit an explicit `passed`, `failed`, or `skipped` result. See `Specs/Testing/Testing_SelfTests.md`.

## CompareDirectories remote smoke (optional)

`RedSalamander.exe --compare-selftest` includes optional remote smoke cases that exercise **cross‑plugin Compare Directories** with Connection Manager profiles:

- `remote_file_s3` (file ↔ S3)
- `remote_file_ftp` (file ↔ FTP)
- `remote_file_sftp` (file ↔ SFTP)
- `remote_file_scp` (file ↔ SCP provider; directory enumeration uses its SSH/SFTP listing route, not an SCP payload transfer)
- `remote_file_imap` (file ↔ IMAP mailbox)
- `remote_file_smb` (local test file ↔ read-only SMB directory comparison)
- `remote_file_onedrive_personal` (file ↔ OneDrive Personal)
- `remote_file_onedrive_business` (file ↔ OneDrive Business)
- `remote_file_sharepoint` (file ↔ SharePoint)
- `remote_s3_pagination` (S3 folder listing pagination: compares non-recursive `GetDirectorySize` vs `ReadDirectoryInfo` count)
- `remote_onedrive_personal_directory_size_callback_contract`
- `remote_onedrive_business_directory_size_callback_contract`
- `remote_sharepoint_directory_size_callback_contract`
- `remote_ftp_continue_on_error_partial` (FTP batch copy/move/rename must return `ERROR_PARTIAL_COPY` when one item succeeds and one item is missing)
- `remote_s3_metadata_smoke` (S3 object metadata + reader path smoke via `GetFileBasicInformation` and `CreateFileReader`)
- `remote_s3_delete_missing` (S3 delete contract: repeated delete returns `ERROR_FILE_NOT_FOUND`; mixed batch delete returns `ERROR_PARTIAL_COPY`)

These cases are **conditional** and use the same profile naming/env var rules as the Phase 16 checks below:

- If the connection profile + secret are present *and* `initialPath` passes the sandbox rules, the case **runs**.
- Otherwise the case is **skipped**.

The compare-only smoke cases are read-only (`remote_file_*`, `remote_s3_pagination`). The FTP/S3 contract cases perform writes and deletes, but only inside a unique per-run child of the configured sandbox root. All remote cases target the remote root as:

- `/@conn:<profileName><initialPath>`

`remote_file_*` reuses one comparison scenario: require a successful remote root
listing, a non-missing root, correct root-relative mapping, and a unique local
fixture reported as OnlyInLeft. It does not create remote fixtures or prove remote
payload transfer, mutation, retry, cancellation, or quota behavior. Record live
smoke results separately from saved-secret/sandbox configuration checks and from
deterministic fake-server fault tests. A declared available resource that skips is
a missing witness to investigate, not successful live coverage.

`remote_file_smb` uses the local filesystem provider directly against
`REDSALAMANDER_SELFTEST_SMB_ROOT`. It is read-only on the SMB side and runs only when
the configured value is an absolute UNC directory below the share, contains no `.` or
`..` segment, and includes a path segment containing `selftest` (case-insensitive).

Note: `remote_s3_pagination` can only *prove* multi-page behavior when the chosen S3 root has **more than 1000 immediate children**
(or when the profile’s S3 `maxKeys` is configured below the expected count).

## Security model

- Secrets are **not** stored in the repo.
- Secrets are stored in **Windows Credential Manager** (WinCred) by Connection Manager when `savePassword == true`.
- Connection profiles (non-secret fields only) live in the app’s Settings Store under `%LOCALAPPDATA%` (outside the source tree).
- OAuth 2.0 PKCE profiles use the host secret kind `refreshToken`; access tokens remain memory-only.

Phase 16 uses `IHostConnections::GetConnectionSecret(...)` and clears the returned memory before freeing it (never writes secrets to logs/artifacts).

## HARD REQUIREMENT: Dedicated remote sandbox root (for remote file-op phases)

Any selftest phase that performs **remote file operations** (copy/move/delete) MUST be sandboxed to a **dedicated, test-only** folder/prefix on the remote side.

- Never run destructive tests against `/`, a home directory, or any user-managed data.
- Remote file-op phases must create a **unique per-run** subfolder under the sandbox root and only touch/delete paths under that per-run folder.

Selftest enforces this requirement via `ConnectionProfile.initialPath` (configuration check):

- Must be an absolute plugin path (starts with `/`) and must not be `/`.
- Must include `selftest` (case-insensitive) to prove the path is test-only.
- For S3, must normally include **bucket + prefix** (at least two path segments), e.g. `/my-test-bucket/red-salamander-selftest`.
- S3 may also use a **bucket root** when the bucket name itself clearly indicates it is test-only, meaning the single bucket segment includes `selftest`, e.g. `/redsalamander-selftest`.

## Setup

1. Open **Connection Manager** in the app.
2. Create a profile for the protocol you want to exercise.
3. Use a stable name (defaults below) or set the env var override.
4. For automated runs:
   - Enable **Save password** (so the secret exists without prompting).
     - For OneDrive/SharePoint this means **remember sign-in** so the refresh token is persisted.
   - Use a dedicated profile whose **Require Windows Hello** policy permits the
     explicitly authorized automated run. Prefer an isolated test copy; do not
     enable a global bypass as an automated test setup step.

### Governed-run settings isolation

When `REDSALAMANDER_TEST_ROOT` is set, test-enabled Settings Store uses only
`<testRoot>\runs\<runId>\scratch\settings-store\RedSalamander\Settings`.
It does not fall back to the normal `%LOCALAPPDATA%` settings. Profile-name
environment variables and the machine resource manifest select profiles; neither
imports their definitions into this isolated store. Thus a profile visible in the
normal app is not, by itself, a configured live-test prerequisite.

For an explicitly requested configured-server run:

1. Reserve a fresh run ID using the existing TestSandbox helpers, beneath an
   initialized, ownership-marked fixed-drive test root.
2. Seed only the selected, reviewed non-secret profile definitions into that run's
   isolated settings, using the current `schemaVersion` and preserving stable
   profile IDs so WinCred lookup remains unchanged. Use
   `RedSalamander-debug.settings.json` for Debug or the current versioned settings
   filename for Release. Do not copy the complete user settings or legacy
   plaintext plugin credentials. Review provider-specific `extra` fields before
   copying them; do not relax authentication, host-key, TLS, or Windows Hello policy
   without explicit scope-specific authorization. An approved Hello exception may
   set `requireWindowsHello=false` only on the named isolated profile copies and
   exact approved roots. Keep the global bypass off, preserve normal settings,
   verify their before/after digest, and record the approval scope with the run.
3. Invoke the governed runner with the same run ID and exact intended cases.
   Keep the normal settings untouched and never archive the seeded settings file.
4. Inspect every selected case and skip reason. Saved-secret and sandbox checks
   prove configuration only; only a successfully executed live case proves its
   stated remote behavior.

Saved-secret access still requires the existing initialized host window and
UI-thread/message-dispatch lifetime, even when Windows Hello is disabled. An
`ERROR_INVALID_WINDOW_HANDLE` result is a host-fixture prerequisite failure, not
proof that credentials or the server are unavailable. Host-dependent test setup
provides that owner through `CompareDirectoriesSelfTest::RequiresHostServices`;
pure engine and UNC-only selections remain headless. Never work around a missing
owner by reading WinCred outside the host or weakening the production guard.

Windows Hello and transport trust are independent policies. In automation,
`HostServices::BuildConnectionJsonUtf8` rejects an insecure-TLS profile with
`ERROR_ACCESS_DENIED` unless both the existing automation allowance and the
profile's TLS acknowledgement are present. A Hello-only test exception grants
neither. Record the refusal as missing live coverage; do not silently change TLS
settings or reinterpret an access-denied comparison as a successful server check.

Compare supports comma-separated exact case names, for example
`remote_file_ftp,remote_file_sftp,remote_file_scp,remote_file_imap,remote_file_s3`.
FileOps currently accepts a single case, family, or prefix, not a comma-separated
set. Select its `Phase16_Remote*Secret`/`Sandbox` checks individually when mutation
is not intended: a broad Phase 16 filter also selects live S3/OneDrive CRUD.

### Default profile names

- FTP: `FileOpsSelfTest FTP`
- SFTP: `FileOpsSelfTest SFTP`
- SCP: `FileOpsSelfTest SCP`
- IMAP: `FileOpsSelfTest IMAP`
- S3: `FileOpsSelfTest S3`
- S3 alternate target: `FileOpsSelfTest S3 Alt`
- OneDrive Personal: `FileOpsSelfTest OneDrive Personal`
- OneDrive Business: `FileOpsSelfTest OneDrive Business`
- SharePoint: `FileOpsSelfTest SharePoint`

### Environment variable overrides (names only; not secrets)

- `REDSALAMANDER_SELFTEST_CONN_FTP`
- `REDSALAMANDER_SELFTEST_CONN_SFTP`
- `REDSALAMANDER_SELFTEST_CONN_SCP`
- `REDSALAMANDER_SELFTEST_CONN_IMAP`
- `REDSALAMANDER_SELFTEST_CONN_S3`
- `REDSALAMANDER_SELFTEST_CONN_S3_ALT`
- `REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_PERSONAL`
- `REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_BUSINESS`
- `REDSALAMANDER_SELFTEST_CONN_SHAREPOINT`
- `REDSALAMANDER_SELFTEST_SMB_ROOT`
- `REDSALAMANDER_TEST_ALTERNATE_ROOT`

`REDSALAMANDER_SELFTEST_CONN_S3_ALT` is used by FileOperations phases that need
a second S3 profile to verify connection override and cross-profile behavior.
It follows the same sandbox-root rules as the primary S3 profile.

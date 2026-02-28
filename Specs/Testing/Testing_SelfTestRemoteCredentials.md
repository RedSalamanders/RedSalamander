# SelfTest Remote Credentials (Optional)

`RedSalamander.exe --fileops-selftest` includes **Phase 16** checks that validate the host can retrieve **saved secrets** (passwords / SSH key passphrases) for remote filesystem plugins:

- FTP (`builtin/file-system-ftp`)
- SFTP (`builtin/file-system-sftp`)
- SCP (`builtin/file-system-scp`)
- IMAP (`builtin/file-system-imap`)
- S3 (`builtin/file-system-s3`)

These phases are **conditional**:

- If the connection profile and secret are present, the phase **passes**.
- Otherwise the phase is **skipped** (selftests stay green by default).

## Security model

- Secrets are **not** stored in the repo.
- Secrets are stored in **Windows Credential Manager** (WinCred) by Connection Manager when `savePassword == true`.
- Connection profiles (non-secret fields only) live in the app’s Settings Store under `%LOCALAPPDATA%` (outside the source tree).

Phase 16 uses `IHostConnections::GetConnectionSecret(...)` and clears the returned memory before freeing it (never writes secrets to logs/artifacts).

## HARD REQUIREMENT: Dedicated remote sandbox root (for remote file-op phases)

Any selftest phase that performs **remote file operations** (copy/move/delete) MUST be sandboxed to a **dedicated, test-only** folder/prefix on the remote side.

- Never run destructive tests against `/`, a home directory, or any user-managed data.
- Remote file-op phases must create a **unique per-run** subfolder under the sandbox root and only touch/delete paths under that per-run folder.

Selftest enforces this requirement via `ConnectionProfile.initialPath` (configuration check):

- Must be an absolute plugin path (starts with `/`) and must not be `/`.
- Must include `selftest` (case-insensitive) to prove the path is test-only.
- For S3, must include **bucket + prefix** (at least two path segments), e.g. `/my-test-bucket/red-salamander-selftest`.

## Setup

1. Open **Connection Manager** in the app.
2. Create a profile for the protocol you want to exercise.
3. Use a stable name (defaults below) or set the env var override.
4. For automated runs:
   - Enable **Save password** (so the secret exists without prompting).
   - Disable **Require Windows Hello** for that profile, *or* enable global `bypassWindowsHello` in settings.

### Default profile names

- FTP: `FileOpsSelfTest FTP`
- SFTP: `FileOpsSelfTest SFTP`
- SCP: `FileOpsSelfTest SCP`
- IMAP: `FileOpsSelfTest IMAP`
- S3: `FileOpsSelfTest S3`

### Environment variable overrides (names only; not secrets)

- `REDSALAMANDER_SELFTEST_CONN_FTP`
- `REDSALAMANDER_SELFTEST_CONN_SFTP`
- `REDSALAMANDER_SELFTEST_CONN_SCP`
- `REDSALAMANDER_SELFTEST_CONN_IMAP`
- `REDSALAMANDER_SELFTEST_CONN_S3`

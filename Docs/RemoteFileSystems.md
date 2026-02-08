# Remote File Systems (FTP / SFTP / SCP / IMAP)

RedSalamander includes four remote virtual file systems:

- `ftp:` (FTP)
- `sftp:` (SFTP)
- `scp:` (SCP)
- `imap:` (IMAP)

They are implemented by a single plugin DLL: `Plugins/FileSystemCurl/FileSystemCurl.dll`.

## Recommended: use Connection Manager (secure secrets)

Instead of embedding passwords in URIs or storing defaults in plain text plugin settings, prefer the host-owned **Connection Manager**:

- Open it from **Commands → Connections Manager…**
- Navigate to a saved connection by name:
  - `nav:<connectionName>`
  - `@conn:<connectionName>`

See: [Connections](Connections.md)

## Navigation syntax (address bar)

You can type standard URI-style locations:

- FTP: `ftp://[user-info@]host[:port][/path]`
- SFTP: `sftp://[user[:password]@]host[:port][/path]`
- SCP: `scp://[user@]host[:port][/path]` *(no password in URI)*
- IMAP: `imap://[user[:password]@]host[:port][/mailbox][/message.eml]`

Short form (uses plugin defaults; requires `defaultHost`):

- `scheme:/path`

Connection Manager authority shorthand (routes to `/<scheme>:/@conn:<name>/...`):

- `ftp://@conn/<name>/...`
- `sftp://@conn/<name>/...`
- `scp://@conn/<name>/...`
- `imap://@conn/<name>/...`

If you type `ftp:` / `sftp:` / `scp:` / `imap:` with no host/path, RedSalamander opens Connection Manager filtered to that protocol.

Examples:

- `sftp://user@example.com/home/user/`
- `imap://user@example.com/INBOX/`
- `ftp:/pub/` *(after configuring plugin defaults)*
- `nav:MySftpServer` *(saved connection profile)*
- `sftp://@conn/MySftpServer/var/log/`

## Plugin settings (plain-text defaults)

Open **View → Preferences… → Plugins** → select the desired file system (FTP / SFTP / SCP / IMAP) to configure defaults like:

- `defaultHost`, `defaultPort`, `defaultUser`, `defaultPassword`, `defaultBasePath`
- Timeouts: `connectTimeoutMs`, `operationTimeoutMs`
- FTP-only: `ftpUseEpsv`
- SFTP/SCP: SSH key paths (`sshPrivateKey`, `sshPublicKey`), passphrase, `sshKnownHosts`
- IMAP: `ignoreSslTrust` and mailbox prefix via `defaultBasePath`

Security note:

- These plugin defaults (including passwords/passphrases) are stored as **plain text** in settings. Prefer [Connection Manager](Connections.md) for secure secrets storage.

## Protocol notes and behavior

- Copy/move between different endpoints uses download → upload via a local temporary file.
- SCP directory operations require the server to support SFTP over SSH.
- IMAP mailbox hierarchy is reconstructed using the IMAP `LIST` delimiter; UI separators are always `/`.
- IMAP messages are exposed as `.eml` files inside mailboxes:
  - Listing names are formatted as `<subject>｜<from>｜<uid>.eml`.
  - `<uid>.eml` is accepted for direct navigation; deleting an `.eml` deletes the message.
- IMAP TLS behavior:
  - Port `993` uses implicit TLS.
  - Other ports use IMAP with STARTTLS when available.

For deep implementation details, see the specs:

- `Specs/FileSystemFtpSftpScpPluginSpec.md`
- `Specs/FileSystemImapPluginSpec.md`


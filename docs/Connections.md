# Connections (Connection Manager)

Connection Manager is the recommended way to store and use remote connection profiles for:

- FTP / SFTP / SCP / IMAP
- Google Drive / OneDrive / SharePoint
- S3 / S3 Table

It keeps non-secret fields in settings and stores secrets using Windows facilities (Credential Manager, with optional Windows Hello gating).

![Connection Manager dialog](res/connections-manager.png)

## Open Connection Manager

- **Commands → Connections Manager…**
- Type `nav:`, `nav://`, or `@conn:` in the address bar (empty name opens the dialog)
- Type a protocol with no host (examples: `sftp:`, `gdrive:`, `onedrivep:`, `sharepoint:`, `s3:`, `s3table:`) to open the dialog filtered to that protocol

## Create and use a profile

1. Open Connection Manager.
2. Create a new connection (or edit an existing one).
3. Set the target (host/port/initial path) and authentication.
4. Choose whether to **Save password**.
5. Press **Connect**.

Then you can navigate to it later by name:

- `nav:MyServer`
- `nav://MyServer`
- `@conn:MyServer`
- `s3://@conn/MyAwsS3/logs/`
- `s3table://@conn/MyDataCatalog/default/`
- `gdrive:/@conn:MyGoogleDrive/`
- `onedrivep:/@conn:MyOneDrive/`
- `sharepoint:/@conn:MySharePointSite/Shared Documents/`

### Advanced path forms

The same saved profile can also be addressed through a protocol-local path:

- `sftp:/@conn:MyServer/var/log/`
- `s3:/@conn:MyAwsS3/logs/`
- `s3://@conn/MyAwsS3/logs/`
- `gdrive:/@conn:MyGoogleDrive/`
- `onedrivep:/@conn:MyOneDrive/`
- `onedriveb:/@conn:MyWorkDrive/Documents/`
- `sharepoint://@conn/MySharePointSite/Shared Documents/`

For cloud-drive setup details and current limitations, see: [Cloud Drives](CloudDrives.md)

### Quick Connect

Connection Manager exposes a `<Quick Connect>` entry that is **not persisted**. It is useful for one-off connections during the current app run.

During the current app run, that temporary profile can also be reopened by name with `@conn:@quick`.

Unlike a saved profile, Quick Connect keeps any secret you enter **in memory only** for the current app run — it is never written to Windows Credential Manager and is discarded when you close the app. Use a saved profile when you want secrets stored durably (and optionally gated by Windows Hello).

## Security notes

- Saving a secret stores it via Windows Credential Manager.
- When **Require Windows Hello** is enabled, Windows Hello verification is performed before secrets are released.
- Some file-system plugins still have “defaults” in Preferences; those values may be stored as plain text. Prefer Connection Manager when possible.

Two global settings under `connections` in the settings file tune Windows Hello behavior across all profiles:

- `bypassWindowsHello` (default `false`): when `true`, Windows Hello verification is skipped even when a profile requires it. This is intended for automation; leave it disabled for normal interactive use.
- `windowsHelloReauthTimeoutMinute` (default `10`): how long a recent successful authentication (Windows Hello verification or a manually entered password/passphrase) is reused for the same connection before you are asked again. Set it to `0` to always prompt on every secret access.

## External settings changes

- If the main settings file changes on disk while Connection Manager is open and you have no unsaved edits, the dialog reloads its non-secret profile data automatically.
- If you have unsaved edits, the dialog prompts **Reload from disk** or **Keep editing**.
- If you keep editing, the next save-producing action asks whether to **Overwrite disk**, **Reload from disk**, or **Cancel**.
- Secrets stored in Credential Manager are not rewritten by this reload path.

## Not implemented yet

- Import/export connections
- Full SSH host-key management UI

# Cloud Drives (Google Drive / OneDrive / SharePoint)

RedSalamander can expose cloud storage through Connection Manager, but the current experience is different for Google Drive and Microsoft-backed drives.

## What works today

| Service | Can sign in from the UI? | Main operations | Notes |
|---------|---------------------------|-----------------|-------|
| Google Drive | No | Browse folders, inspect drive info | Current milestone is read-only metadata only. No download, upload, rename, move, or delete yet. |
| OneDrive Personal | Yes | Browse, read, write, rename, move, delete, directory size | Uses Microsoft Graph with browser sign-in. |
| OneDrive Business | Yes | Browse, read, write, rename, move, delete, directory size | Uses Microsoft Graph with browser sign-in. |
| SharePoint | Yes | Browse, read, write, rename, move, delete, directory size | Same plugin family as OneDrive Business, but rooted to a SharePoint site or library. |

## Before you create a connection

Open **Plugins -> Plugin Manager...** and configure the relevant plugin first.

### Google Drive

Set the Google Drive plugin's `defaultClientId` unless every connection will provide its own client id through advanced settings.

Important limitation:

- The current Google Drive plugin does **not** launch the first-time browser sign-in flow.
- A Google Drive connection only works if a refresh token has already been stored for that connection by host-side OAuth plumbing or test tooling.
- If no refresh token exists, connection attempts fail with an authentication error.

### OneDrive Personal / OneDrive Business / SharePoint

RedSalamander ships with a built-in Microsoft Graph `clientId` for the Microsoft drive plugins.

You only need to open plugin configuration if you want to override that default app registration.

Important limitation:

- If the plugin `clientId` is cleared or replaced with an invalid value, Microsoft sign-in cannot start.

## Create a connection profile

1. Open **Commands -> Connections Manager...**
2. Create a new connection.
3. Pick the protocol:
   - `Google Drive`
   - `OneDrive Personal`
   - `OneDrive Business`
   - `SharePoint`
4. Enter a unique **Name**. This is the name you will later use with `nav:MyConnection`.
5. Set **Initial path**:
   - Usually `/`
   - Or a subfolder such as `/Documents/Reports/`
6. Set **Remember sign-in** if you want the refresh token persisted.
7. For SharePoint only, fill **Host** with the tenant host and optional site path:
   - `contoso.sharepoint.com`
   - `contoso.sharepoint.com/sites/Team`
8. Press **Connect**.

### Notes by service

#### Google Drive

- `Host` and `Port` are hidden because Google Drive profiles are hostless.
- The default root is your `My Drive`.
- The current Connection Manager UI does not expose advanced Google-only fields such as shared-drive selection.
- Pressing **Connect** only succeeds if the connection already has a valid refresh token in the host secret store.

#### OneDrive Personal / OneDrive Business

- `Host` and `Port` are hidden because these profiles are hostless.
- `User name` is optional and is used only as a login hint.
- On first connect, RedSalamander opens the system browser and completes OAuth with PKCE.

#### SharePoint

- `Host` is required.
- Use the tenant host or a specific site path.
- `Initial path` is interpreted inside the resolved document library.

## Reopen a saved cloud connection

Once a profile exists, you can open it by name:

- `nav:MyGoogleDrive`
- `nav:MyOneDrive`
- `nav:MySharePointSite`
- `@conn:MyOneDrive`

You can also use protocol-local paths:

- `gdrive:/@conn:MyGoogleDrive/`
- `onedrivep:/@conn:MyOneDrive/`
- `onedriveb:/@conn:MyWorkDrive/Documents/`
- `sharepoint:/@conn:MySharePointSite/Shared Documents/`

Authority shorthand also works:

- `gdrive://@conn/MyGoogleDrive/`
- `onedrivep://@conn/MyOneDrive/`
- `onedriveb://@conn/MyWorkDrive/Documents/`
- `sharepoint://@conn/MySharePointSite/Shared Documents/`

## Advanced settings not exposed in the current dialog

Some cloud-specific options exist in the saved connection profile but are not currently editable in the standard Connection Manager UI.

### Google Drive advanced fields

Stored in the connection `extra` object:

- `rootKind`
  - `myDrive`
  - `sharedDrive`
- `sharedDriveId`
- `googleDocsMode`
- `readOnly`
- `useDefaultClientId`
- `clientId`

Example non-secret profile fields:

```json
{
  "name": "Shared Design Drive",
  "pluginId": "builtin/file-system-gdrive",
  "authMode": "oauth2Pkce",
  "initialPath": "/",
  "savePassword": true,
  "extra": {
    "rootKind": "sharedDrive",
    "sharedDriveId": "0AExampleDriveIdPVA",
    "googleDocsMode": "export",
    "useDefaultClientId": true
  }
}
```

### Microsoft drive advanced fields

Stored in the connection `extra` object:

- `tenantAuthority`
- `driveId`

Example SharePoint profile fields:

```json
{
  "name": "Team Site",
  "pluginId": "builtin/file-system-sharepoint",
  "host": "contoso.sharepoint.com/sites/Team",
  "authMode": "oauth2Pkce",
  "initialPath": "/Shared Documents/",
  "savePassword": true,
  "extra": {
    "tenantAuthority": "organizations"
  }
}
```

Connection Manager secrets are **not** stored in the JSON settings file. Only non-secret profile fields are stored there.

See also:

- [Connections](Connections.md)
- [Navigation & Path Syntax](NavigationAndPaths.md)
- [Settings File & Advanced Configuration](SettingsFile.md)

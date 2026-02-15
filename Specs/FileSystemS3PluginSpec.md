# S3 / S3 Table Virtual File System Plugin

## Overview

RedSalamander ships a built-in virtual file system implementation in `Plugins/FileSystemS3/FileSystemS3.dll` that exposes **two UI-visible file system plugins**:

- **S3** (`s3:`)
- **S3 Table** (`s3table:`)

Both are implemented using the AWS SDK for C++ (`aws-sdk-cpp`), with the S3 CRT client for S3 operations and the S3 Tables client for S3 Table operations.

## Plugin Identities

| Display name | `PluginMetaData.id` | `PluginMetaData.shortId` |
|-------------|----------------------|--------------------------|
| S3          | `builtin/file-system-s3`      | `s3`      |
| S3 Table    | `builtin/file-system-s3table` | `s3table` |

## Navigation (URI Syntax)

### Path normalization rules

- `\` is treated as `/`.
- Multiple slashes are collapsed **except** the leading `//` after the scheme (authority prefix).

### S3

Root behavior:

- `s3:/` lists buckets (as directories).
  - If the active Connection Manager profile has a non-empty region, the list is filtered to buckets in that region.
  - If the region is empty, the list is unfiltered and bucket regions are resolved automatically when entering a bucket.

Within a bucket:

- `s3:/<bucket>/` lists “folders” (common prefixes) and objects under the current prefix.
- Folder traversal is modeled via `ListObjectsV2` with delimiter `/`.

Authority form:

- `s3://<bucket>/<path>` is accepted and is normalized internally to `/<bucket>/<path>`.

Examples:

- `s3:/`
- `s3:/my-bucket/`
- `s3://my-bucket/logs/2026/`

### S3 Table

Root behavior:

- `s3table:/` lists S3 Table buckets (as directories).

Within a table bucket:

- `s3table:/<tableBucket>/` lists namespaces (as directories).
- `s3table:/<tableBucket>/<namespace>/` lists tables as files named `"<table>.table.json"`.

Opening a `*.table.json` “file” returns a JSON document derived from `GetTable` for the selected table.

Examples:

- `s3table:/`
- `s3table:/my-table-bucket/default/`
- `s3table:/my-table-bucket/default/my_table.table.json`

## Configuration (Per Plugin)

Each plugin exposes its own configuration schema via `IInformations::GetConfigurationSchema()` and is configured in Preferences.

### S3 keys

- `defaultRegion` (string, default `us-east-1`)
- `defaultEndpointOverride` (string, default empty)
- `useHttps` (bool, default `true`)
- `verifyTls` (bool, default `true`)
- `useVirtualAddressing` (bool, default `true`)
- `maxKeys` (integer, `1..1000`, default `1000`)

### S3 Table keys

- `defaultRegion` (string, default `us-east-1`)
- `defaultEndpointOverride` (string, default empty)
- `useHttps` (bool, default `true`)
- `verifyTls` (bool, default `true`)
- `maxTableResults` (integer, `1..1000`, default `1000`)

Security note:

- AWS secrets are **not** stored in plugin settings. For secure storage (WinCred + optional Windows Hello), use Connection Manager (`/@conn:<connectionName>/...`).

## Connection Manager Integration

Both plugins support host-reserved navigation:

- `/@conn:<connectionName>/...` (see `Specs/ConnectionManagerSpec.md`)
- Optional shorthand authority form: `// @conn / <connectionName> / ...` (e.g. `s3://@conn/<connectionName>/...`)

Profile mapping:

- `ConnectionProfile.pluginId`:
  - S3: `builtin/file-system-s3`
  - S3 Table: `builtin/file-system-s3table`
- `ConnectionProfile.host`: AWS region (e.g. `us-east-1`).
  - If empty: treated as “auto” (list all buckets, then resolve per-bucket region on demand).
  - If non-empty: used as the signing/endpoint region and as a filter for bucket listing.
- `ConnectionProfile.userName`: Access key id (optional; when omitted the AWS default credential chain is used)
- Secret access key:
  - stored as the connection `password` secret (WinCred generic credential),
  - retrieved via `IHostConnections::GetConnectionSecret(HOST_CONNECTION_SECRET_PASSWORD)` and may be prompted via `PromptForConnectionSecret(...)` when needed.

Supported `extra` keys (from `ConnectionProfile.extra` and surfaced to plugins via `GetConnectionJsonUtf8().extra`):

- `endpointOverride` (string)
- `useHttps` (bool)
- `verifyTls` (bool)
- `useVirtualAddressing` (bool, S3 only)

## Operations and Behavior Notes

- Browsing and file reads are implemented as **read-only** operations.
- Mutating operations (copy/move/delete/rename) currently return `ERROR_NOT_SUPPORTED`.
- File reads download the remote content/metadata to a local delete-on-close temporary file before streaming it to the host.

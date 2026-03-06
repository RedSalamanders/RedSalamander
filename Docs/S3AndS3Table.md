# S3 / S3 Table

RedSalamander includes two AWS-backed virtual file systems implemented by `Plugins/FileSystemS3/FileSystemS3.dll`:

- `s3:` for Amazon S3 object storage
- `s3table:` for S3 Table buckets, namespaces, and tables

## Recommended: use Connection Manager

Use **Commands → Connections Manager…** to store AWS profiles securely.

- Typing `s3:` or `s3table:` with no bucket/path opens Connection Manager filtered to that protocol.
- Saved profiles can be opened with:
  - `nav:<connectionName>`
  - `nav://<connectionName>`
  - `@conn:<connectionName>`
  - `s3://@conn/<connectionName>/...`
  - `s3table://@conn/<connectionName>/...`

Connection Manager stores non-secret fields in settings and keeps secrets in Windows Credential Manager.

## Navigation syntax

### S3

- `s3:/` lists buckets
- `s3:/<bucket>/` lists the current prefix inside a bucket
- `s3://<bucket>/<path>` is also accepted

Examples:

- `s3:/`
- `s3:/my-bucket/`
- `s3://my-bucket/logs/2026/`

### S3 Table

- `s3table:/` lists S3 Table buckets
- `s3table:/<tableBucket>/` lists namespaces
- `s3table:/<tableBucket>/<namespace>/` lists tables as `*.table.json`

Examples:

- `s3table:/`
- `s3table:/my-table-bucket/default/`
- `s3table:/my-table-bucket/default/my_table.table.json`

## What each file system can do

### `s3:`

- Browse buckets, prefixes, and objects
- Open/download objects
- Upload new objects
- Overwrite existing objects
- Delete objects

Current limitations:

- No server-side rename/move
- No recursive delete of virtual folders / prefixes
- "Folders" are virtual path prefixes, not real directory objects

Cross-filesystem copy/move still works through the host bridge (read -> write), so you can copy between `file:`, `ftp:`, `s3:`, and other supported file systems when both sides allow it.

### `s3table:`

- Browse table buckets and namespaces
- Open tables as generated `*.table.json` documents

Current limitations:

- No upload/write
- No delete/rename/move

## Plugin settings

Open **View → Preferences… → Plugins** and select **S3** or **S3 Table** to configure:

- `defaultRegion`
- `defaultEndpointOverride`
- `useHttps`
- `verifyTls`
- `connectTimeoutMs`
- `requestTimeoutMs`

S3 also exposes:

- `useVirtualAddressing`
- `maxKeys`

S3 Table also exposes:

- `maxTableResults`

## Troubleshooting notes

- If a profile keeps prompting for credentials, prefer Connection Manager over embedding defaults in plugin settings.
- If you are using a custom endpoint or self-signed certificate, check the S3/S3 Table plugin settings and your Connection Manager profile.
- If a delete fails on what looks like a folder, verify that you selected objects rather than a virtual prefix.

# Microsoft Drive filesystem plan

Last updated: 2026-03-07

## Goal

Ship a built-in WinHTTP-based Microsoft Graph filesystem plugin covering:

- OneDrive Personal
- OneDrive Business
- SharePoint

The implementation must integrate with Connection Manager, support multiple saved Microsoft identities, and keep dependency growth minimal.

## Chosen implementation

- Transport: WinHTTP
- JSON: `yyjson`
- Resource lifetime: WIL RAII
- Auth: delegated OAuth 2.0 authorization code + PKCE, system browser, localhost loopback redirect
- Secret persistence: host-managed `refreshToken`

## Scope

- One DLL exposing three logical plugins:
  - `builtin/file-system-onedrive-personal`
  - `builtin/file-system-onedrive-business`
  - `builtin/file-system-sharepoint`
- Multiple connection profiles and identities through Connection Manager
- Read, write, mkdir, delete, move, rename, metadata, directory size
- Optional selftests for secret retrieval, sandbox checks, remote compare smoke, and remote directory-size callback contracts

## Status

- [x] Host settings and connection-secret plumbing for `oauth2Pkce`
- [x] Connection Manager UI updates for Microsoft profiles
- [x] New `FileSystemMicrosoftDrive` project and plugin exports
- [x] WinHTTP transport and Graph request pipeline
- [x] PKCE browser sign-in and refresh-token persistence
- [x] OneDrive Personal / Business drive resolution
- [x] SharePoint site + drive resolution
- [x] File IO and directory operations
- [x] Selftest integration
- [x] Core and filesystem specs

## Remaining verification

- [ ] Full `RedSalamander` build in this worktree
- [ ] Focused selftest execution with real Microsoft profiles

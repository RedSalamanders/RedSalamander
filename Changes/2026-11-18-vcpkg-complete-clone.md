- The DxUi restore prepares a complete vcpkg checkout before it runs the pinned DxUi `vcpkg-install.ps1`. That
  installer (pin `7fb2ab71`) clones vcpkg with `--filter=blob:none`, and vcpkg 2026-07-27 extracts each versioned
  port with `git read-tree` and `git checkout-index` in an environment holding only `GIT_INDEX_FILE`, where Git
  resolves no host. Since upstream microsoft/vcpkg de36db0455 (2026-11-18) moved `vcpkg-cmake` past the baseline's
  tree, every fresh restore failed with "Could not resolve host: github.com" for a promisor object. The new
  `Tools/Modules/Build/VcpkgToolCheckout.psm1`, the same file in every DxUi consumer, clones
  `<output>/vcpkg-tool` completely at the pinned source's `vcpkg-tool.json` commit, completes an earlier partial
  or shallow clone in place, refuses another repository and deletes nothing. The installer then reuses that
  checkout. The restore runs the helper in the same contained, logged child process as the installer, and now
  requires `vcpkg-tool.json` in the sparse DxUi checkout. RedSalamander's own vcpkg checkout was already a
  complete clone. `Tools/Tests/VcpkgToolCheckout.Tests.ps1` covers a fresh and a repeated clone, the completed
  partial clone (with the failing state as a negative control), the refusals, an offline extraction of an older
  port tree with lazy fetch disabled, and the restore's own command composition with paths that need quoting.
  The native compile is proved by the post-merge qualification (D21).

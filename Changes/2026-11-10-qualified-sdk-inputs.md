- Reuse optional exact pinned qualified DxUi SDK inputs for Release builds with
  actual compiler and per-translation-unit ABI guards, live reader ownership,
  linker provenance and sealed dependency records. Reopen original SDK inputs
  before artifact reuse and packaging; retain Source fallback and unchanged pins.
  Keep SDK packages on the separate receipt-verified packaging route and reject
  SDK inputs in the existing Source-only portable attestation format.
  Preserve the application's Windows header macro policy while importing SDK
  integration, and retain strict required artifact downloads alongside the exact
  DxUi-only optional library handoff.
  Preserve the verified SDK recipe through artifact-mutating tooling and profile
  re-attestation; block lost, substituted or changed SDK input before native tests.
  Resolve production localization/plugin targets through the full solution-folder graph, preserving required projects and rejecting invalid target metadata before compilation.
  Keep SDK CI builds and test reuse on canonical vcpkg MSBuild paths without injecting a compiler override that the ABI guard must reject.

  Alternate-volume test acquisition now refuses uninitialized drive-root folders;
  the generic opt-in allows use only of an existing, explicitly initialized marked
  sandbox. It no longer creates a new root or ownership marker during discovery.
  Give each factory-created native child a private compiler metadata-service endpoint so one owned operation cannot retire another operation's shared service.

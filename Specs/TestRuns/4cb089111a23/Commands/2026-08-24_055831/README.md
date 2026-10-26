# File Operations P3.1 configuration/migration validation

The focused Debug x64 Commands run passed `settings_file_system_plugin_roundtrip` 1/1. It proves the
File System plugin schema exposes only Preserve/Skip, legacy Follow and Copy-Reparse values serialize
to their canonical replacements, unknown/numeric policy values are rejected, and reserved ABI policy
zero fails before provider I/O. This is focused configuration evidence, not a Full repository verdict
or a performance-improvement claim.

Run id: `20260824T035810Z-55356-15f34e4518ad45d8996f45e1bcdde25a`.

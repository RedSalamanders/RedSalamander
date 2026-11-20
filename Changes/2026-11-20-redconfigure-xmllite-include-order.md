- Restored the SDK include order for RedConfigure and its tests on x64 and ARM64.
  `objidl.h` now precedes `XmlLite.h` in a separate include block so clang-format
  preserves the COM declarations that `XmlLite.h` needs.

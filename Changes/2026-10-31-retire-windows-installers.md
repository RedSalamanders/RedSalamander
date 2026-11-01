- Remove MSI and MSIX packaging from the RedSalamander solution, build command,
  release automation and active documentation. Keep portable ZIP and winget
  delivery with receipt validation and clean-extraction smoke checks; reject
  ProductionOnly plus ZIP because that build omits the matching smoke-test host.

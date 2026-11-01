- Preserve parallel search cancellation and callback failures when the last result
  arrives after workers finish; cover that ordering with a deterministic fixture
  gate and keep normal successful completion unchanged.

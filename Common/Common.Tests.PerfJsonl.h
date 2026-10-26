#pragma once

#ifdef ENABLE_TESTS
#include "Helpers.h"

namespace Debug::Perf
{
struct JsonlDebugSnapshot
{
    uint64_t threadWriteCalls     = 0u;
    uint64_t threadWriteElapsedUs = 0u;
    uint64_t metadataPreparations = 0u;
    size_t retainedLineCapacity   = 0u;
};

// Observation only: this thread's cumulative inclusive writer elapsed time, plus
// locked shared resource counters. It never writes a metric or changes capture.
COMMON_API JsonlDebugSnapshot GetJsonlDebugSnapshotForSelfTest() noexcept;
} // namespace Debug::Perf
#endif

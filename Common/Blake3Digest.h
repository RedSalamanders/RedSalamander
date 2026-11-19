#pragma once

#include "Common/Common.h"

#include <blake3.h>

#include <array>
#include <cstddef>
#include <span>

namespace Common::Crypto
{
inline constexpr size_t kBlake3DigestBytes = BLAKE3_OUT_LEN;
using Blake3Digest                         = std::array<std::byte, kBlake3DigestBytes>;

// Canonical streaming BLAKE3 primitive. File Operations owns I/O, cancellation, progress, and
// exact-object authority; this helper owns only byte-order-preserving digest state.
class COMMON_API Blake3Hasher final
{
public:
    Blake3Hasher() noexcept;

    void Update(std::span<const std::byte> bytes) noexcept;
    [[nodiscard]] Blake3Digest Finalize() const noexcept;

private:
    blake3_hasher _hasher{};
};
} // namespace Common::Crypto

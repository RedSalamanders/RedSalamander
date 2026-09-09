#include "Blake3Digest.h"

namespace Common::Crypto
{
Blake3Hasher::Blake3Hasher() noexcept
{
    blake3_hasher_init(&_hasher);
}

void Blake3Hasher::Update(std::span<const std::byte> bytes) noexcept
{
    if (! bytes.empty())
    {
        blake3_hasher_update(&_hasher, bytes.data(), bytes.size());
    }
}

Blake3Digest Blake3Hasher::Finalize() const noexcept
{
    Blake3Digest digest{};
    blake3_hasher_finalize(&_hasher, reinterpret_cast<uint8_t*>(digest.data()), digest.size());
    return digest;
}
} // namespace Common::Crypto

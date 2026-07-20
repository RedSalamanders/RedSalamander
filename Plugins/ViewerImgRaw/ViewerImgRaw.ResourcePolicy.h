#pragma once

#include <cstdint>
#include <limits>

namespace ViewerImgRawResource
{
inline constexpr uint64_t kMebibyte = 1024ull * 1024ull;

struct DecodedImagePolicy final
{
    uint32_t maxDimension         = 16'384u;
    uint64_t maxPixels            = 64ull * 1024ull * 1024ull;
    uint64_t maxBgraBytes         = 256ull * kMebibyte;
    uint64_t maxEmbeddedJpegBytes = 64ull * kMebibyte;
};

inline constexpr DecodedImagePolicy kProductionDecodedImagePolicy{};
inline constexpr uint64_t kProductionSpeculativeDecodedBytes = 384ull * kMebibyte;
inline constexpr uint64_t kProductionMainSourceBytes         = 1024ull * kMebibyte;
inline constexpr uint64_t kProductionSidecarSourceBytes      = 64ull * kMebibyte;
inline constexpr uint64_t kProductionMainOwnedPeakBytes =
    kProductionMainSourceBytes + kProductionSidecarSourceBytes + (2ull * kProductionDecodedImagePolicy.maxBgraBytes);

enum class ValidationError : uint8_t
{
    None = 0,
    InvalidDimensions,
    DimensionLimit,
    PixelLimit,
    ByteLimit,
    Overflow,
    InvalidFormat,
    SourceLength,
};

struct DecodedImageLayout final
{
    uint32_t width     = 0u;
    uint32_t height    = 0u;
    uint64_t pixels    = 0u;
    uint64_t rowBytes  = 0u;
    uint64_t bgraBytes = 0u;
};

[[nodiscard]] constexpr bool TryMultiply(uint64_t lhs, uint64_t rhs, uint64_t& out) noexcept
{
    out = 0u;
    if (lhs == 0u || rhs == 0u)
    {
        return true;
    }
    if (lhs > (std::numeric_limits<uint64_t>::max)() / rhs)
    {
        return false;
    }
    out = lhs * rhs;
    return true;
}

[[nodiscard]] constexpr bool IsProviderReadCountValid(unsigned long requestedBytes, unsigned long returnedBytes) noexcept
{
    return returnedBytes <= requestedBytes;
}

[[nodiscard]] constexpr ValidationError ValidateDecodedImage(uint32_t width,
                                                             uint32_t height,
                                                             const DecodedImagePolicy& policy,
                                                             DecodedImageLayout& out) noexcept
{
    out = {};
    if (width == 0u || height == 0u || policy.maxDimension == 0u || policy.maxPixels == 0u || policy.maxBgraBytes == 0u)
    {
        return ValidationError::InvalidDimensions;
    }
    if (width > policy.maxDimension || height > policy.maxDimension)
    {
        return ValidationError::DimensionLimit;
    }

    uint64_t pixels = 0u;
    if (! TryMultiply(width, height, pixels))
    {
        return ValidationError::Overflow;
    }
    if (pixels == 0u || pixels > policy.maxPixels)
    {
        return ValidationError::PixelLimit;
    }

    uint64_t rowBytes  = 0u;
    uint64_t bgraBytes = 0u;
    if (! TryMultiply(width, 4u, rowBytes) || ! TryMultiply(pixels, 4u, bgraBytes))
    {
        return ValidationError::Overflow;
    }
    if (bgraBytes > policy.maxBgraBytes)
    {
        return ValidationError::ByteLimit;
    }

    out.width     = width;
    out.height    = height;
    out.pixels    = pixels;
    out.rowBytes  = rowBytes;
    out.bgraBytes = bgraBytes;
    return ValidationError::None;
}

[[nodiscard]] constexpr ValidationError ValidateEmbeddedJpeg(
    uint32_t width, uint32_t height, uint64_t compressedBytes, uint64_t sourceBytes, const DecodedImagePolicy& policy, DecodedImageLayout& out) noexcept
{
    const ValidationError imageError = ValidateDecodedImage(width, height, policy, out);
    if (imageError != ValidationError::None)
    {
        return imageError;
    }
    if (compressedBytes == 0u || compressedBytes > sourceBytes)
    {
        return ValidationError::SourceLength;
    }
    if (compressedBytes > policy.maxEmbeddedJpegBytes)
    {
        return ValidationError::ByteLimit;
    }
    return ValidationError::None;
}

[[nodiscard]] constexpr ValidationError ValidatePackedBitmap(uint32_t width,
                                                             uint32_t height,
                                                             uint32_t colors,
                                                             uint32_t bitsPerSample,
                                                             uint64_t packedBytes,
                                                             uint64_t sourceBytes,
                                                             const DecodedImagePolicy& policy,
                                                             DecodedImageLayout& out,
                                                             uint64_t& outExpectedPackedBytes) noexcept
{
    outExpectedPackedBytes           = 0u;
    const ValidationError imageError = ValidateDecodedImage(width, height, policy, out);
    if (imageError != ValidationError::None)
    {
        return imageError;
    }
    if (colors == 0u || colors > 4u || (bitsPerSample != 8u && bitsPerSample != 16u))
    {
        return ValidationError::InvalidFormat;
    }
    if (packedBytes == 0u || packedBytes > sourceBytes)
    {
        return ValidationError::SourceLength;
    }

    uint64_t sampleCount = 0u;
    if (! TryMultiply(out.pixels, colors, sampleCount) || ! TryMultiply(sampleCount, bitsPerSample / 8u, outExpectedPackedBytes))
    {
        return ValidationError::Overflow;
    }
    if (outExpectedPackedBytes > packedBytes || outExpectedPackedBytes > sourceBytes)
    {
        return ValidationError::SourceLength;
    }
    return ValidationError::None;
}
} // namespace ViewerImgRawResource

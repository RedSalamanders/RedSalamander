#pragma once

// Header-only streaming digests for provider content proofs (R3-2). The host hashes the bytes it
// streams into a provider writer with the algorithm that provider can prove after publication
// (SHA-256 / SHA-1 / MD5 through CNG, Microsoft's QuickXorHash, or the CRC64-NVME checksum S3
// returns for a whole object); the provider returns the digest it holds for the published object
// and the bridge compares the two. This is content-integrity evidence for verification and Managed
// Move source cleanup; it never authorizes overwrite, skip, or deletion by itself.

#include <windows.h>
#include <bcrypt.h>

#include <array>
#include <cstddef>
#include <cstdint>
#include <span>
#include <string>
#include <string_view>
#include <vector>

#pragma comment(lib, "bcrypt.lib")

namespace Common::Crypto
{
enum class ContentDigestAlgorithm : uint32_t
{
    None      = 0,
    Sha256    = 2, // matches FILESYSTEM_CONTENT_PROOF_SHA256_256
    Sha1      = 3, // FILESYSTEM_CONTENT_PROOF_SHA1_160
    Md5       = 4, // FILESYSTEM_CONTENT_PROOF_MD5_128
    QuickXor  = 5, // FILESYSTEM_CONTENT_PROOF_QUICKXOR_160
    Crc64Nvme = 6, // FILESYSTEM_CONTENT_PROOF_CRC64NVME_64
};

[[nodiscard]] inline constexpr size_t ContentDigestBytes(ContentDigestAlgorithm algorithm) noexcept
{
    switch (algorithm)
    {
        case ContentDigestAlgorithm::Sha256: return 32u;
        case ContentDigestAlgorithm::Sha1: return 20u;
        case ContentDigestAlgorithm::Md5: return 16u;
        case ContentDigestAlgorithm::QuickXor: return 20u;
        case ContentDigestAlgorithm::Crc64Nvme: return 8u;
        case ContentDigestAlgorithm::None:
        default: return 0u;
    }
}

namespace Detail
{
// CRC-64/NVME (reflected, polynomial 0xAD93D23594C93659, init and xorout all ones); the CRC S3
// reports as x-amz-checksum-crc64nvme. Check value for "123456789" is 0xAE8B14860A799888.
inline const std::array<uint64_t, 256>& Crc64NvmeTable() noexcept
{
    static const std::array<uint64_t, 256> table = []() noexcept
    {
        constexpr uint64_t kReflectedPolynomial = 0x9A6C9329AC4BC9B5ull; // bit-reflected 0xAD93D23594C93659
        std::array<uint64_t, 256> t{};
        for (uint32_t i = 0; i < 256u; ++i)
        {
            uint64_t crc = i;
            for (int bit = 0; bit < 8; ++bit)
            {
                crc = (crc & 1u) != 0u ? (crc >> 1) ^ kReflectedPolynomial : crc >> 1;
            }
            t[i] = crc;
        }
        return t;
    }();
    return table;
}

class CngHasher final
{
public:
    CngHasher() = default;
    ~CngHasher()
    {
        Reset();
    }
    CngHasher(const CngHasher&)            = delete;
    CngHasher& operator=(const CngHasher&) = delete;
    CngHasher(CngHasher&&)                 = delete;
    CngHasher& operator=(CngHasher&&)      = delete;

    [[nodiscard]] bool Open(const wchar_t* algorithmId, size_t expectedBytes) noexcept
    {
        Reset();
        BCRYPT_ALG_HANDLE algorithm = nullptr;
        if (! BCRYPT_SUCCESS(BCryptOpenAlgorithmProvider(&algorithm, algorithmId, nullptr, 0)) || algorithm == nullptr)
        {
            return false;
        }
        _algorithm     = algorithm;
        DWORD objectLength = 0;
        DWORD hashLength   = 0;
        DWORD written      = 0;
        if (! BCRYPT_SUCCESS(BCryptGetProperty(_algorithm, BCRYPT_OBJECT_LENGTH, reinterpret_cast<PUCHAR>(&objectLength), sizeof(objectLength), &written, 0)) ||
            objectLength == 0 ||
            ! BCRYPT_SUCCESS(BCryptGetProperty(_algorithm, BCRYPT_HASH_LENGTH, reinterpret_cast<PUCHAR>(&hashLength), sizeof(hashLength), &written, 0)) ||
            hashLength != expectedBytes)
        {
            Reset();
            return false;
        }
        // Allocation failure is fatal under the repository policy. This method is noexcept, so
        // letting vector assignment terminate is intentional; false is reserved for CNG errors.
        _object.assign(objectLength, std::byte{});
        BCRYPT_HASH_HANDLE hash = nullptr;
        if (! BCRYPT_SUCCESS(BCryptCreateHash(_algorithm, &hash, reinterpret_cast<PUCHAR>(_object.data()), objectLength, nullptr, 0, 0)) || hash == nullptr)
        {
            Reset();
            return false;
        }
        _hash      = hash;
        _hashBytes = hashLength;
        return true;
    }

    [[nodiscard]] bool Update(std::span<const std::byte> bytes) noexcept
    {
        if (_hash == nullptr)
        {
            return false;
        }
        size_t offset = 0;
        while (offset < bytes.size())
        {
            const size_t chunk = (std::min)(bytes.size() - offset, static_cast<size_t>(1u << 30));
            if (! BCRYPT_SUCCESS(BCryptHashData(_hash, const_cast<PUCHAR>(reinterpret_cast<const UCHAR*>(bytes.data() + offset)), static_cast<ULONG>(chunk), 0)))
            {
                return false;
            }
            offset += chunk;
        }
        return true;
    }

    [[nodiscard]] bool Finish(std::vector<std::byte>& digest) noexcept
    {
        if (_hash == nullptr)
        {
            return false;
        }
        // Allocation failure is fatal; false remains an honest CNG-operation result.
        digest.assign(_hashBytes, std::byte{});
        const bool ok = BCRYPT_SUCCESS(BCryptFinishHash(_hash, reinterpret_cast<PUCHAR>(digest.data()), _hashBytes, 0));
        Reset();
        return ok;
    }

private:
    void Reset() noexcept
    {
        if (_hash != nullptr)
        {
            BCryptDestroyHash(_hash);
            _hash = nullptr;
        }
        if (_algorithm != nullptr)
        {
            BCryptCloseAlgorithmProvider(_algorithm, 0);
            _algorithm = nullptr;
        }
        _object.clear();
        _hashBytes = 0;
    }

    BCRYPT_ALG_HANDLE _algorithm = nullptr;
    BCRYPT_HASH_HANDLE _hash     = nullptr;
    std::vector<std::byte> _object;
    DWORD _hashBytes = 0;
};

// Microsoft QuickXorHash (OneDrive `file.hashes.quickXorHash`): a 160-bit rotating XOR where each
// input byte is shifted 11 bits further than the previous one, with the total length XORed into
// the last 8 bytes. Port of the published reference algorithm.
class QuickXorHasher final
{
public:
    void Update(std::span<const std::byte> bytes) noexcept
    {
        constexpr int kWidthInBits     = 160;
        constexpr int kShift           = 11;
        constexpr int kBitsInLastCell  = 32;
        int currentShift               = _shiftSoFar;
        int vectorArrayIndex           = currentShift / 64;
        int vectorOffset               = currentShift % 64;
        const size_t iterations        = (std::min)(bytes.size(), static_cast<size_t>(kWidthInBits));
        for (size_t i = 0; i < iterations; ++i)
        {
            const bool isLastCell     = vectorArrayIndex == static_cast<int>(_data.size()) - 1;
            const int bitsInVectorCell = isLastCell ? kBitsInLastCell : 64;
            if (vectorOffset <= bitsInVectorCell - 8)
            {
                for (size_t j = i; j < bytes.size(); j += kWidthInBits)
                {
                    _data[static_cast<size_t>(vectorArrayIndex)] ^= static_cast<uint64_t>(bytes[j]) << vectorOffset;
                }
            }
            else
            {
                const int index1 = vectorArrayIndex;
                const int index2 = isLastCell ? 0 : vectorArrayIndex + 1;
                const int low    = bitsInVectorCell - vectorOffset;
                for (size_t j = i; j < bytes.size(); j += kWidthInBits)
                {
                    const uint64_t value = static_cast<uint64_t>(bytes[j]);
                    _data[static_cast<size_t>(index1)] ^= value << vectorOffset;
                    _data[static_cast<size_t>(index2)] ^= value >> low;
                }
            }
            vectorOffset += kShift;
            while (vectorOffset >= bitsInVectorCell)
            {
                vectorArrayIndex = isLastCell ? 0 : vectorArrayIndex + 1;
                vectorOffset -= bitsInVectorCell;
            }
        }
        _shiftSoFar = static_cast<int>((static_cast<uint64_t>(_shiftSoFar) + static_cast<uint64_t>(kShift) * (bytes.size() % kWidthInBits)) % kWidthInBits);
        _lengthSoFar += static_cast<uint64_t>(bytes.size());
    }

    [[nodiscard]] std::array<std::byte, 20> Finalize() const noexcept
    {
        std::array<std::byte, 20> digest{};
        // Cells are laid out little-endian: two full 64-bit cells and the low 32 bits of the third.
        for (size_t i = 0; i < 8u; ++i)
        {
            digest[i]      = static_cast<std::byte>((_data[0] >> (8u * i)) & 0xFFu);
            digest[8u + i] = static_cast<std::byte>((_data[1] >> (8u * i)) & 0xFFu);
        }
        for (size_t i = 0; i < 4u; ++i)
        {
            digest[16u + i] = static_cast<std::byte>((_data[2] >> (8u * i)) & 0xFFu);
        }
        for (size_t i = 0; i < 8u; ++i)
        {
            digest[12u + i] ^= static_cast<std::byte>((_lengthSoFar >> (8u * i)) & 0xFFu);
        }
        return digest;
    }

private:
    std::array<uint64_t, 3> _data{};
    int _shiftSoFar        = 0;
    uint64_t _lengthSoFar  = 0;
};
} // namespace Detail

// One streaming hasher per algorithm. `Update` and `Finish` report false only for CNG failures;
// a hasher that could not be opened stays invalid and every call reports false.
class ContentHasher final
{
public:
    explicit ContentHasher(ContentDigestAlgorithm algorithm) noexcept : _algorithm(algorithm)
    {
        switch (algorithm)
        {
            case ContentDigestAlgorithm::Sha256: _valid = _cng.Open(BCRYPT_SHA256_ALGORITHM, 32u); break;
            case ContentDigestAlgorithm::Sha1: _valid = _cng.Open(BCRYPT_SHA1_ALGORITHM, 20u); break;
            case ContentDigestAlgorithm::Md5: _valid = _cng.Open(BCRYPT_MD5_ALGORITHM, 16u); break;
            case ContentDigestAlgorithm::QuickXor:
            case ContentDigestAlgorithm::Crc64Nvme: _valid = true; break;
            case ContentDigestAlgorithm::None:
            default: _valid = false; break;
        }
    }
    ContentHasher(const ContentHasher&)            = delete;
    ContentHasher& operator=(const ContentHasher&) = delete;
    ContentHasher(ContentHasher&&)                 = delete;
    ContentHasher& operator=(ContentHasher&&)      = delete;
    ~ContentHasher()                               = default;

    [[nodiscard]] bool Valid() const noexcept
    {
        return _valid;
    }
    [[nodiscard]] ContentDigestAlgorithm Algorithm() const noexcept
    {
        return _algorithm;
    }

    [[nodiscard]] bool Update(std::span<const std::byte> bytes) noexcept
    {
        if (! _valid)
        {
            return false;
        }
        switch (_algorithm)
        {
            case ContentDigestAlgorithm::Sha256:
            case ContentDigestAlgorithm::Sha1:
            case ContentDigestAlgorithm::Md5: _valid = _cng.Update(bytes); return _valid;
            case ContentDigestAlgorithm::QuickXor: _quickXor.Update(bytes); return true;
            case ContentDigestAlgorithm::Crc64Nvme:
            {
                const auto& table = Detail::Crc64NvmeTable();
                uint64_t crc      = _crc;
                for (const std::byte value : bytes)
                {
                    crc = table[static_cast<size_t>((crc ^ static_cast<uint64_t>(value)) & 0xFFu)] ^ (crc >> 8);
                }
                _crc = crc;
                return true;
            }
            case ContentDigestAlgorithm::None:
            default: return false;
        }
    }

    // Produces the digest in the byte order providers report (big-endian CRC, raw hash bytes).
    [[nodiscard]] bool Finish(std::vector<std::byte>& digest) noexcept
    {
        digest.clear();
        if (! _valid)
        {
            return false;
        }
        switch (_algorithm)
        {
            case ContentDigestAlgorithm::Sha256:
            case ContentDigestAlgorithm::Sha1:
            case ContentDigestAlgorithm::Md5: _valid = _cng.Finish(digest); return _valid;
            case ContentDigestAlgorithm::QuickXor:
            {
                const std::array<std::byte, 20> value = _quickXor.Finalize();
                digest.assign(value.begin(), value.end());
                return true;
            }
            case ContentDigestAlgorithm::Crc64Nvme:
            {
                const uint64_t value = _crc ^ 0xFFFFFFFFFFFFFFFFull;
                digest.resize(8u);
                for (size_t i = 0; i < 8u; ++i)
                {
                    digest[i] = static_cast<std::byte>((value >> (8u * (7u - i))) & 0xFFu);
                }
                return true;
            }
            case ContentDigestAlgorithm::None:
            default: return false;
        }
    }

private:
    ContentDigestAlgorithm _algorithm = ContentDigestAlgorithm::None;
    bool _valid                       = false;
    Detail::CngHasher _cng;
    Detail::QuickXorHasher _quickXor;
    uint64_t _crc = 0xFFFFFFFFFFFFFFFFull;
};

// One-shot digest of a buffer (fixtures and small provider objects).
[[nodiscard]] inline bool ComputeContentDigest(ContentDigestAlgorithm algorithm, std::span<const std::byte> bytes, std::vector<std::byte>& digest) noexcept
{
    ContentHasher hasher(algorithm);
    return hasher.Update(bytes) && hasher.Finish(digest);
}

[[nodiscard]] inline bool DecodeHexDigest(std::string_view text, std::vector<std::byte>& digest) noexcept
{
    digest.clear();
    if (text.empty() || (text.size() % 2u) != 0u)
    {
        return false;
    }
    const auto nibble = [](char ch) noexcept -> int
    {
        if (ch >= '0' && ch <= '9') return ch - '0';
        if (ch >= 'a' && ch <= 'f') return ch - 'a' + 10;
        if (ch >= 'A' && ch <= 'F') return ch - 'A' + 10;
        return -1;
    };
    digest.reserve(text.size() / 2u);
    for (size_t i = 0; i < text.size(); i += 2u)
    {
        const int high = nibble(text[i]);
        const int low  = nibble(text[i + 1u]);
        if (high < 0 || low < 0)
        {
            digest.clear();
            return false;
        }
        digest.push_back(static_cast<std::byte>((high << 4) | low));
    }
    return true;
}

[[nodiscard]] inline bool DecodeBase64Digest(std::string_view text, std::vector<std::byte>& digest) noexcept
{
    digest.clear();
    const auto value = [](char ch) noexcept -> int
    {
        if (ch >= 'A' && ch <= 'Z') return ch - 'A';
        if (ch >= 'a' && ch <= 'z') return ch - 'a' + 26;
        if (ch >= '0' && ch <= '9') return ch - '0' + 52;
        if (ch == '+') return 62;
        if (ch == '/') return 63;
        return -1;
    };
    uint32_t accumulator = 0;
    int bits             = 0;
    for (const char ch : text)
    {
        if (ch == '=')
        {
            break;
        }
        const int v = value(ch);
        if (v < 0)
        {
            digest.clear();
            return false;
        }
        accumulator = (accumulator << 6) | static_cast<uint32_t>(v);
        bits += 6;
        if (bits >= 8)
        {
            bits -= 8;
            digest.push_back(static_cast<std::byte>((accumulator >> bits) & 0xFFu));
        }
    }
    return ! digest.empty();
}

[[nodiscard]] inline std::string EncodeBase64Digest(std::span<const std::byte> digest)
{
    static constexpr char kAlphabet[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    std::string out;
    out.reserve(((digest.size() + 2u) / 3u) * 4u);
    size_t i = 0;
    while (i + 3u <= digest.size())
    {
        const uint32_t triple = (static_cast<uint32_t>(digest[i]) << 16) | (static_cast<uint32_t>(digest[i + 1u]) << 8) | static_cast<uint32_t>(digest[i + 2u]);
        out.push_back(kAlphabet[(triple >> 18) & 0x3Fu]);
        out.push_back(kAlphabet[(triple >> 12) & 0x3Fu]);
        out.push_back(kAlphabet[(triple >> 6) & 0x3Fu]);
        out.push_back(kAlphabet[triple & 0x3Fu]);
        i += 3u;
    }
    if (i + 1u == digest.size())
    {
        const uint32_t value = static_cast<uint32_t>(digest[i]) << 16;
        out.push_back(kAlphabet[(value >> 18) & 0x3Fu]);
        out.push_back(kAlphabet[(value >> 12) & 0x3Fu]);
        out += "==";
    }
    else if (i + 2u == digest.size())
    {
        const uint32_t value = (static_cast<uint32_t>(digest[i]) << 16) | (static_cast<uint32_t>(digest[i + 1u]) << 8);
        out.push_back(kAlphabet[(value >> 18) & 0x3Fu]);
        out.push_back(kAlphabet[(value >> 12) & 0x3Fu]);
        out.push_back(kAlphabet[(value >> 6) & 0x3Fu]);
        out.push_back('=');
    }
    return out;
}

[[nodiscard]] inline std::string EncodeHexDigest(std::span<const std::byte> digest)
{
    static constexpr char kHex[] = "0123456789abcdef";
    std::string out;
    out.reserve(digest.size() * 2u);
    for (const std::byte value : digest)
    {
        out.push_back(kHex[(static_cast<unsigned>(value) >> 4) & 0xFu]);
        out.push_back(kHex[static_cast<unsigned>(value) & 0xFu]);
    }
    return out;
}

// Known-answer checks (used by debug self-tests): CRC-64/NVME of "123456789" and SHA-256 of "abc".
[[nodiscard]] inline bool ContentDigestSelfCheck() noexcept
{
    static constexpr char kCheck[] = "123456789";
    std::vector<std::byte> crc;
    if (! ComputeContentDigest(ContentDigestAlgorithm::Crc64Nvme, std::as_bytes(std::span<const char>(kCheck, 9u)), crc) || crc.size() != 8u)
    {
        return false;
    }
    static constexpr std::array<std::byte, 8> kCrcExpected{
        {std::byte{0xAE}, std::byte{0x8B}, std::byte{0x14}, std::byte{0x86}, std::byte{0x0A}, std::byte{0x79}, std::byte{0x98}, std::byte{0x88}}};
    if (! std::equal(crc.begin(), crc.end(), kCrcExpected.begin()))
    {
        return false;
    }
    static constexpr char kAbc[] = "abc";
    std::vector<std::byte> sha;
    if (! ComputeContentDigest(ContentDigestAlgorithm::Sha256, std::as_bytes(std::span<const char>(kAbc, 3u)), sha) || sha.size() != 32u ||
        EncodeHexDigest(sha) != "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad")
    {
        return false;
    }
    std::vector<std::byte> sha1;
    if (! ComputeContentDigest(ContentDigestAlgorithm::Sha1, std::as_bytes(std::span<const char>(kAbc, 3u)), sha1) || sha1.size() != 20u ||
        EncodeHexDigest(sha1) != "a9993e364706816aba3e25717850c26c9cd0d89d")
    {
        return false;
    }
    // QuickXorHash of "abc" from an independent transcription of Microsoft's published algorithm.
    std::vector<std::byte> quickXor;
    return ComputeContentDigest(ContentDigestAlgorithm::QuickXor, std::as_bytes(std::span<const char>(kAbc, 3u)), quickXor) && quickXor.size() == 20u &&
           EncodeHexDigest(quickXor) == "6110c31800000000000000000300000000000000";
}
} // namespace Common::Crypto

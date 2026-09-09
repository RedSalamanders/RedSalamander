#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <wil/resource.h>

#include "TerminalEngineGate0Adapter.h"
#include <TerminalEngineFixtures.v1.h>

#include <algorithm>
#include <array>
#include <atomic>
#include <bit>
#include <charconv>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <limits>
#include <mutex>
#include <ranges>
#include <span>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

namespace
{
namespace FixtureContract = RedSalamander::TerminalEngine::Fixtures::V1;

constexpr uint16_t kColumns              = 120;
constexpr uint16_t kRows                 = 40;
constexpr uint64_t kRequiredCapabilities = RS_TERMINAL_GATE0_CAP_CELLS_AND_STYLES | RS_TERMINAL_GATE0_CAP_SNAPSHOT_COPY | RS_TERMINAL_GATE0_CAP_INPUT_ENCODING |
                                           RS_TERMINAL_GATE0_CAP_EFFECT_ORDERING | RS_TERMINAL_GATE0_CAP_KITTY_STATIC |
                                           RS_TERMINAL_GATE0_CAP_RESOURCE_ACCOUNTING | RS_TERMINAL_GATE0_CAP_QUIET_DESTROY;
constexpr size_t kMaximumFixtureRows     = 64;
constexpr size_t kMaximumSessions        = 32;

#if defined(_M_X64)
constexpr RsTerminalGate0Architecture kExpectedArchitecture = RS_TERMINAL_GATE0_ARCH_X64;
constexpr std::string_view kArchitectureName                = "x64";
#elif defined(_M_ARM64)
constexpr RsTerminalGate0Architecture kExpectedArchitecture = RS_TERMINAL_GATE0_ARCH_ARM64;
constexpr std::string_view kArchitectureName                = "ARM64";
#else
#error The terminal Gate-0 harness supports only x64 and ARM64.
#endif

struct Arguments
{
    std::string adapterDll;
    std::string expectedCandidateId;
    std::string expectedUpstreamPin;
    std::string expectedSourceSha256;
    std::string expectedAdapterHeaderSha256;
    std::string expectedBuildIdentity;
    std::string expectedFixtureManifestSha256;
    uint64_t expectedCapabilities         = 0;
    uint32_t expectedAdapterAbi           = 0;
    uint64_t addedReleasePackageBytes     = 0;
    uint64_t crossedAbiElements           = 0;
    uint64_t nonSystemRuntimeDependencies = 0;
    bool fullMeasurements                 = false;
    bool hasAddedReleasePackageBytes      = false;
    bool hasCrossedAbiElements            = false;
    bool hasNonSystemRuntimeDependencies  = false;
};

struct Sha256
{
    [[nodiscard]] bool Update(std::span<const uint8_t> bytes) noexcept
    {
        if (_finished || bytes.size() > std::numeric_limits<uint64_t>::max() - _totalBytes)
        {
            return false;
        }
        _totalBytes += static_cast<uint64_t>(bytes.size());
        size_t offset = 0;
        if (_bufferLength != 0)
        {
            const size_t copied = std::min(bytes.size(), _buffer.size() - _bufferLength);
            std::memcpy(_buffer.data() + _bufferLength, bytes.data(), copied);
            _bufferLength += copied;
            offset += copied;
            if (_bufferLength == _buffer.size())
            {
                Transform(_buffer.data());
                _bufferLength = 0;
            }
        }
        while (bytes.size() - offset >= _buffer.size())
        {
            Transform(bytes.data() + offset);
            offset += _buffer.size();
        }
        if (offset != bytes.size())
        {
            _bufferLength = bytes.size() - offset;
            std::memcpy(_buffer.data(), bytes.data() + offset, _bufferLength);
        }
        return true;
    }

    [[nodiscard]] bool Finish(std::array<uint8_t, 32>& digest) noexcept
    {
        if (_finished || _totalBytes > std::numeric_limits<uint64_t>::max() / 8)
        {
            return false;
        }
        const uint64_t bitLength = _totalBytes * 8;
        _buffer[_bufferLength++] = 0x80;
        if (_bufferLength > 56)
        {
            std::fill(_buffer.begin() + static_cast<ptrdiff_t>(_bufferLength), _buffer.end(), uint8_t{0});
            Transform(_buffer.data());
            _bufferLength = 0;
        }
        std::fill(_buffer.begin() + static_cast<ptrdiff_t>(_bufferLength), _buffer.begin() + 56, uint8_t{0});
        for (size_t index = 0; index < 8; ++index)
        {
            _buffer[63 - index] = static_cast<uint8_t>(bitLength >> (index * 8));
        }
        Transform(_buffer.data());
        for (size_t index = 0; index < _state.size(); ++index)
        {
            digest[index * 4]     = static_cast<uint8_t>(_state[index] >> 24);
            digest[index * 4 + 1] = static_cast<uint8_t>(_state[index] >> 16);
            digest[index * 4 + 2] = static_cast<uint8_t>(_state[index] >> 8);
            digest[index * 4 + 3] = static_cast<uint8_t>(_state[index]);
        }
        _finished = true;
        return true;
    }

private:
    [[nodiscard]] static constexpr uint32_t RotateRight(uint32_t value, unsigned count) noexcept
    {
        return std::rotr(value, static_cast<int>(count));
    }

    void Transform(const uint8_t* block) noexcept
    {
        static constexpr std::array<uint32_t, 64> constants = {
            0x428a2f98U, 0x71374491U, 0xb5c0fbcfU, 0xe9b5dba5U, 0x3956c25bU, 0x59f111f1U, 0x923f82a4U, 0xab1c5ed5U, 0xd807aa98U, 0x12835b01U, 0x243185beU,
            0x550c7dc3U, 0x72be5d74U, 0x80deb1feU, 0x9bdc06a7U, 0xc19bf174U, 0xe49b69c1U, 0xefbe4786U, 0x0fc19dc6U, 0x240ca1ccU, 0x2de92c6fU, 0x4a7484aaU,
            0x5cb0a9dcU, 0x76f988daU, 0x983e5152U, 0xa831c66dU, 0xb00327c8U, 0xbf597fc7U, 0xc6e00bf3U, 0xd5a79147U, 0x06ca6351U, 0x14292967U, 0x27b70a85U,
            0x2e1b2138U, 0x4d2c6dfcU, 0x53380d13U, 0x650a7354U, 0x766a0abbU, 0x81c2c92eU, 0x92722c85U, 0xa2bfe8a1U, 0xa81a664bU, 0xc24b8b70U, 0xc76c51a3U,
            0xd192e819U, 0xd6990624U, 0xf40e3585U, 0x106aa070U, 0x19a4c116U, 0x1e376c08U, 0x2748774cU, 0x34b0bcb5U, 0x391c0cb3U, 0x4ed8aa4aU, 0x5b9cca4fU,
            0x682e6ff3U, 0x748f82eeU, 0x78a5636fU, 0x84c87814U, 0x8cc70208U, 0x90befffaU, 0xa4506cebU, 0xbef9a3f7U, 0xc67178f2U,
        };
        std::array<uint32_t, 64> words{};
        for (size_t index = 0; index < 16; ++index)
        {
            const size_t offset = index * 4;
            words[index]        = (static_cast<uint32_t>(block[offset]) << 24) | (static_cast<uint32_t>(block[offset + 1]) << 16) |
                                  (static_cast<uint32_t>(block[offset + 2]) << 8) | static_cast<uint32_t>(block[offset + 3]);
        }
        for (size_t index = 16; index < words.size(); ++index)
        {
            const uint32_t s0 = RotateRight(words[index - 15], 7) ^ RotateRight(words[index - 15], 18) ^ (words[index - 15] >> 3);
            const uint32_t s1 = RotateRight(words[index - 2], 17) ^ RotateRight(words[index - 2], 19) ^ (words[index - 2] >> 10);
            words[index]      = words[index - 16] + s0 + words[index - 7] + s1;
        }
        uint32_t a = _state[0];
        uint32_t b = _state[1];
        uint32_t c = _state[2];
        uint32_t d = _state[3];
        uint32_t e = _state[4];
        uint32_t f = _state[5];
        uint32_t g = _state[6];
        uint32_t h = _state[7];
        for (size_t index = 0; index < words.size(); ++index)
        {
            const uint32_t sum1       = RotateRight(e, 6) ^ RotateRight(e, 11) ^ RotateRight(e, 25);
            const uint32_t choose     = (e & f) ^ (~e & g);
            const uint32_t temporary1 = h + sum1 + choose + constants[index] + words[index];
            const uint32_t sum0       = RotateRight(a, 2) ^ RotateRight(a, 13) ^ RotateRight(a, 22);
            const uint32_t majority   = (a & b) ^ (a & c) ^ (b & c);
            const uint32_t temporary2 = sum0 + majority;
            h                         = g;
            g                         = f;
            f                         = e;
            e                         = d + temporary1;
            d                         = c;
            c                         = b;
            b                         = a;
            a                         = temporary1 + temporary2;
        }
        _state[0] += a;
        _state[1] += b;
        _state[2] += c;
        _state[3] += d;
        _state[4] += e;
        _state[5] += f;
        _state[6] += g;
        _state[7] += h;
    }

    std::array<uint32_t, 8> _state = {
        0x6a09e667U,
        0xbb67ae85U,
        0x3c6ef372U,
        0xa54ff53aU,
        0x510e527fU,
        0x9b05688cU,
        0x1f83d9abU,
        0x5be0cd19U,
    };
    std::array<uint8_t, 64> _buffer{};
    size_t _bufferLength = 0;
    uint64_t _totalBytes = 0;
    bool _finished       = false;
};

[[nodiscard]] std::string DigestHex(const std::array<uint8_t, 32>& digest)
{
    static constexpr std::string_view hex = "0123456789abcdef";
    std::string result(digest.size() * 2, '\0');
    for (size_t index = 0; index < digest.size(); ++index)
    {
        result[index * 2]     = hex[digest[index] >> 4];
        result[index * 2 + 1] = hex[digest[index] & 0x0f];
    }
    return result;
}

[[nodiscard]] bool Sha256Hex(std::span<const uint8_t> bytes, std::string& result)
{
    Sha256 hash;
    std::array<uint8_t, 32> digest{};
    if (! hash.Update(bytes) || ! hash.Finish(digest))
    {
        return false;
    }
    result = DigestHex(digest);
    return true;
}

[[nodiscard]] bool Sha256Hex(std::string_view bytes, std::string& result)
{
    return Sha256Hex(std::span(reinterpret_cast<const uint8_t*>(bytes.data()), bytes.size()), result);
}

[[nodiscard]] bool IsLowerHex(std::string_view value, size_t length) noexcept
{
    return value.size() == length &&
           std::ranges::all_of(value, [](char character) { return (character >= '0' && character <= '9') || (character >= 'a' && character <= 'f'); });
}

[[nodiscard]] bool ParseUnsigned(std::string_view value, uint64_t& output, int base) noexcept
{
    if (value.empty())
    {
        return false;
    }
    if (base == 16 && (value.starts_with("0x") || value.starts_with("0X")))
    {
        value.remove_prefix(2);
    }
    const auto parsed = std::from_chars(value.data(), value.data() + value.size(), output, base);
    return parsed.ec == std::errc{} && parsed.ptr == value.data() + value.size();
}

[[nodiscard]] bool ParseArguments(int argc, char** argv, Arguments& arguments) noexcept
{
    auto consume = [&](int& index, std::string_view option) -> std::string_view
    {
        if (argv[index] == option && index + 1 < argc)
        {
            ++index;
            return argv[index];
        }
        return {};
    };
    bool adapterAbi      = false;
    bool capabilities    = false;
    bool measurementMode = false;
    for (int index = 1; index < argc; ++index)
    {
        std::string_view value;
        if (! (value = consume(index, "--adapter-dll")).empty())
        {
            arguments.adapterDll = value;
        }
        else if (! (value = consume(index, "--expected-candidate-id")).empty())
        {
            arguments.expectedCandidateId = value;
        }
        else if (! (value = consume(index, "--expected-upstream-pin")).empty())
        {
            arguments.expectedUpstreamPin = value;
        }
        else if (! (value = consume(index, "--expected-source-sha256")).empty())
        {
            arguments.expectedSourceSha256 = value;
        }
        else if (! (value = consume(index, "--expected-adapter-header-sha256")).empty())
        {
            arguments.expectedAdapterHeaderSha256 = value;
        }
        else if (! (value = consume(index, "--expected-build-identity")).empty())
        {
            arguments.expectedBuildIdentity = value;
        }
        else if (! (value = consume(index, "--expected-fixture-manifest-sha256")).empty())
        {
            arguments.expectedFixtureManifestSha256 = value;
        }
        else if (! (value = consume(index, "--expected-adapter-abi")).empty())
        {
            uint64_t parsed = 0;
            if (! ParseUnsigned(value, parsed, 10) || parsed > std::numeric_limits<uint32_t>::max())
            {
                return false;
            }
            arguments.expectedAdapterAbi = static_cast<uint32_t>(parsed);
            adapterAbi                   = true;
        }
        else if (! (value = consume(index, "--expected-capabilities")).empty())
        {
            if (! ParseUnsigned(value, arguments.expectedCapabilities, 16))
            {
                return false;
            }
            capabilities = true;
        }
        else if (! (value = consume(index, "--measurement-mode")).empty())
        {
            if (value == "functional")
            {
                arguments.fullMeasurements = false;
            }
            else if (value == "full")
            {
                arguments.fullMeasurements = true;
            }
            else
            {
                return false;
            }
            measurementMode = true;
        }
        else if (! (value = consume(index, "--added-release-package-bytes")).empty())
        {
            if (! ParseUnsigned(value, arguments.addedReleasePackageBytes, 10))
            {
                return false;
            }
            arguments.hasAddedReleasePackageBytes = true;
        }
        else if (! (value = consume(index, "--crossed-abi-elements")).empty())
        {
            if (! ParseUnsigned(value, arguments.crossedAbiElements, 10))
            {
                return false;
            }
            arguments.hasCrossedAbiElements = true;
        }
        else if (! (value = consume(index, "--non-system-runtime-dependencies")).empty())
        {
            if (! ParseUnsigned(value, arguments.nonSystemRuntimeDependencies, 10))
            {
                return false;
            }
            arguments.hasNonSystemRuntimeDependencies = true;
        }
        else
        {
            return false;
        }
    }
    return ! arguments.adapterDll.empty() && ! arguments.expectedCandidateId.empty() && ! arguments.expectedUpstreamPin.empty() &&
           ! arguments.expectedSourceSha256.empty() && ! arguments.expectedAdapterHeaderSha256.empty() && ! arguments.expectedBuildIdentity.empty() &&
           ! arguments.expectedFixtureManifestSha256.empty() && adapterAbi && capabilities && measurementMode &&
           (! arguments.fullMeasurements ||
            (arguments.hasAddedReleasePackageBytes && arguments.hasCrossedAbiElements && arguments.hasNonSystemRuntimeDependencies)) &&
           arguments.expectedAdapterAbi == RS_TERMINAL_GATE0_ADAPTER_ABI_V1 && IsLowerHex(arguments.expectedSourceSha256, 64) &&
           IsLowerHex(arguments.expectedAdapterHeaderSha256, 64) && IsLowerHex(arguments.expectedFixtureManifestSha256, 64);
}

[[nodiscard]] int HexNibble(char character) noexcept
{
    if (character >= '0' && character <= '9')
    {
        return character - '0';
    }
    if (character >= 'a' && character <= 'f')
    {
        return character - 'a' + 10;
    }
    return -1;
}

[[nodiscard]] bool ExpandFixtureInput(const FixtureContract::Fixture& fixture, std::vector<uint8_t>& output)
{
    if (fixture.inputSizeBytes > static_cast<uint64_t>(std::numeric_limits<size_t>::max()))
    {
        return false;
    }
    output.clear();
    output.reserve(static_cast<size_t>(fixture.inputSizeBytes));
    for (const auto& step : fixture.inputPlan)
    {
        switch (step.op)
        {
            case FixtureContract::InputOp::LiteralHex:
                if ((step.hex.size() & 1U) != 0)
                {
                    return false;
                }
                for (size_t index = 0; index < step.hex.size(); index += 2)
                {
                    const int high = HexNibble(step.hex[index]);
                    const int low  = HexNibble(step.hex[index + 1]);
                    if (high < 0 || low < 0)
                    {
                        return false;
                    }
                    output.push_back(static_cast<uint8_t>((static_cast<unsigned>(high) << 4) | static_cast<unsigned>(low)));
                }
                break;
            case FixtureContract::InputOp::RepeatByte:
                if (step.repeatCount > static_cast<uint64_t>(std::numeric_limits<size_t>::max() - output.size()))
                {
                    return false;
                }
                output.insert(output.end(), static_cast<size_t>(step.repeatCount), step.byteValue);
                break;
            default: return false;
        }
    }
    std::string sha256;
    return output.size() == fixture.inputSizeBytes && Sha256Hex(output, sha256) && sha256 == fixture.inputSha256;
}

[[nodiscard]] const FixtureContract::Check* FindCheck(const FixtureContract::Fixture& fixture, std::string_view id) noexcept
{
    const auto found = std::ranges::find(fixture.checks, id, &FixtureContract::Check::checkId);
    return found == fixture.checks.end() ? nullptr : &*found;
}

void AppendJcsString(std::string& output, std::string_view value)
{
    static constexpr std::string_view hex = "0123456789abcdef";
    output.push_back('"');
    for (const unsigned char character : value)
    {
        switch (character)
        {
            case '"': output.append("\\\""); break;
            case '\\': output.append("\\\\"); break;
            case '\b': output.append("\\b"); break;
            case '\f': output.append("\\f"); break;
            case '\n': output.append("\\n"); break;
            case '\r': output.append("\\r"); break;
            case '\t': output.append("\\t"); break;
            default:
                if (character < 0x20)
                {
                    output.append("\\u00");
                    output.push_back(hex[character >> 4]);
                    output.push_back(hex[character & 0x0f]);
                }
                else
                {
                    output.push_back(static_cast<char>(character));
                }
                break;
        }
    }
    output.push_back('"');
}

[[nodiscard]] bool AppendDecimalString(std::string& output, uint64_t value)
{
    std::array<char, 32> buffer{};
    const auto converted = std::to_chars(buffer.data(), buffer.data() + buffer.size(), value);
    if (converted.ec != std::errc{})
    {
        return false;
    }
    output.push_back('"');
    output.append(buffer.data(), static_cast<size_t>(converted.ptr - buffer.data()));
    output.push_back('"');
    return true;
}

[[nodiscard]] std::string_view ScheduleAlgorithmName(FixtureContract::ScheduleAlgorithm algorithm) noexcept
{
    switch (algorithm)
    {
        case FixtureContract::ScheduleAlgorithm::AllAtOnce: return "AllAtOnce";
        case FixtureContract::ScheduleAlgorithm::Bytewise: return "Bytewise";
        case FixtureContract::ScheduleAlgorithm::Fibonacci: return "Fibonacci";
        case FixtureContract::ScheduleAlgorithm::TerminatorCuts: return "TerminatorCuts";
    }
    return {};
}

[[nodiscard]] std::string_view ExpectedScheduleId(FixtureContract::ScheduleAlgorithm algorithm) noexcept
{
    switch (algorithm)
    {
        case FixtureContract::ScheduleAlgorithm::AllAtOnce: return "all-at-once";
        case FixtureContract::ScheduleAlgorithm::Bytewise: return "bytewise";
        case FixtureContract::ScheduleAlgorithm::Fibonacci: return "fibonacci";
        case FixtureContract::ScheduleAlgorithm::TerminatorCuts: return "terminator-split";
    }
    return {};
}

[[nodiscard]] bool ValidateScheduleIdentity(const FixtureContract::Fixture& fixture, const FixtureContract::ScheduleSpec& schedule)
{
    const std::string_view algorithm = ScheduleAlgorithmName(schedule.algorithm);
    if (algorithm.empty() || schedule.scheduleId != ExpectedScheduleId(schedule.algorithm) || ! IsLowerHex(schedule.scheduleSha256, 64) ||
        (schedule.algorithm != FixtureContract::ScheduleAlgorithm::TerminatorCuts && ! schedule.cutOffsets.empty()))
    {
        return false;
    }
    uint64_t prior = 0;
    for (const uint64_t cut : schedule.cutOffsets)
    {
        if (cut == 0 || cut >= fixture.inputSizeBytes || cut <= prior)
        {
            return false;
        }
        prior = cut;
    }
    std::string canonical;
    canonical.append("{\"algorithm\":");
    AppendJcsString(canonical, algorithm);
    canonical.append(",\"cutOffsets\":[");
    for (size_t index = 0; index < schedule.cutOffsets.size(); ++index)
    {
        if (index != 0)
        {
            canonical.push_back(',');
        }
        if (! AppendDecimalString(canonical, schedule.cutOffsets[index]))
        {
            return false;
        }
    }
    canonical.append("],\"inputSizeBytes\":");
    if (! AppendDecimalString(canonical, fixture.inputSizeBytes))
    {
        return false;
    }
    canonical.append(",\"scheduleId\":");
    AppendJcsString(canonical, schedule.scheduleId);
    canonical.push_back('}');
    std::string sha256;
    return Sha256Hex(canonical, sha256) && sha256 == schedule.scheduleSha256;
}

[[nodiscard]] bool ValidateFixtureIdentity(const FixtureContract::Fixture& fixture, std::string& contractSha256)
{
    std::vector<uint8_t> input;
    if (! ExpandFixtureInput(fixture, input) || ! IsLowerHex(fixture.contractSha256, 64))
    {
        return false;
    }
    std::string canonical;
    canonical.append("{\"actions\":[");
    for (size_t index = 0; index < fixture.actions.size(); ++index)
    {
        const auto& action = fixture.actions[index];
        std::string actionCanonical;
        actionCanonical.append("{\"actionId\":");
        AppendJcsString(actionCanonical, action.actionId);
        actionCanonical.append(",\"arguments\":");
        actionCanonical.append(action.argumentsJcs);
        actionCanonical.push_back('}');
        std::string actionSha256;
        if (! Sha256Hex(actionCanonical, actionSha256) || actionSha256 != action.actionSha256)
        {
            return false;
        }
        if (index != 0)
        {
            canonical.push_back(',');
        }
        canonical.append("{\"actionId\":");
        AppendJcsString(canonical, action.actionId);
        canonical.append(",\"actionSha256\":");
        AppendJcsString(canonical, action.actionSha256);
        canonical.push_back('}');
    }
    canonical.append("],\"checks\":[");
    for (size_t index = 0; index < fixture.checks.size(); ++index)
    {
        const auto& check = fixture.checks[index];
        std::string expectedSha256;
        if (! IsLowerHex(check.expectedSha256, 64) || ! Sha256Hex(check.expectedValue, expectedSha256) || expectedSha256 != check.expectedSha256)
        {
            return false;
        }
        if (index != 0)
        {
            canonical.push_back(',');
        }
        canonical.append("{\"checkId\":");
        AppendJcsString(canonical, check.checkId);
        canonical.append(",\"expectedSha256\":");
        AppendJcsString(canonical, check.expectedSha256);
        canonical.push_back('}');
    }
    canonical.append("],\"fixtureId\":");
    AppendJcsString(canonical, fixture.fixtureId);
    canonical.append(",\"inputSha256\":");
    AppendJcsString(canonical, fixture.inputSha256);
    canonical.append(",\"inputSizeBytes\":");
    if (! AppendDecimalString(canonical, fixture.inputSizeBytes))
    {
        return false;
    }
    canonical.append(",\"schedules\":[");
    for (size_t index = 0; index < fixture.schedules.size(); ++index)
    {
        const auto& schedule = fixture.schedules[index];
        if (! ValidateScheduleIdentity(fixture, schedule))
        {
            return false;
        }
        if (index != 0)
        {
            canonical.push_back(',');
        }
        canonical.append("{\"scheduleId\":");
        AppendJcsString(canonical, schedule.scheduleId);
        canonical.append(",\"scheduleSha256\":");
        AppendJcsString(canonical, schedule.scheduleSha256);
        canonical.push_back('}');
    }
    canonical.append("]}");
    return Sha256Hex(canonical, contractSha256) && contractSha256 == fixture.contractSha256;
}

[[nodiscard]] std::vector<uint8_t> BuildPerformancePayload();

[[nodiscard]] bool VerifyFrozenFixtureShape(const Arguments& arguments) noexcept
{
    if (FixtureContract::kFixtureManifestSha256 != arguments.expectedFixtureManifestSha256 || FixtureContract::GetFixtures().size() != 29 ||
        FixtureContract::kPerformancePayload.sizeBytes != 8U * 1024U * 1024U || ! IsLowerHex(FixtureContract::kPerformancePayload.sha256, 64))
    {
        return false;
    }
    std::string emptySha256;
    std::string abcSha256;
    if (! Sha256Hex(std::string_view{}, emptySha256) || ! Sha256Hex("abc", abcSha256) ||
        emptySha256 != "e3b0c44298fc1c149afbf4c8996fb924"
                       "27ae41e4649b934ca495991b7852b855" ||
        abcSha256 != "ba7816bf8f01cfea414140de5dae2223"
                     "b00361a396177a9cb410ff61f20015ad")
    {
        return false;
    }
    size_t scheduleCount = 0;
    std::string manifestCanonical("[");
    std::string_view priorFixture;
    size_t fixtureIndex = 0;
    for (const auto& fixture : FixtureContract::GetFixtures())
    {
        if (fixture.fixtureId.empty() || (! priorFixture.empty() && priorFixture.compare(fixture.fixtureId) >= 0) || fixture.actions.empty() ||
            fixture.actions.front().actionId != "feed-current-schedule" ||
            fixture.actions.front().argumentsJcs != "{\"input\":\"inputPlan\",\"schedule\":\"current\"}" || fixture.checks.size() != 2 ||
            fixture.checks[0].checkId != "resource-trace" || fixture.checks[1].checkId != "state" || FindCheck(fixture, "resource-trace") == nullptr ||
            FindCheck(fixture, "state") == nullptr)
        {
            return false;
        }
        for (size_t index = 0; index < fixture.actions.size(); ++index)
        {
            const auto& action = fixture.actions[index];
            if (action.actionId.empty() || action.argumentsJcs.empty() || ! IsLowerHex(action.actionSha256, 64))
            {
                return false;
            }
            for (size_t prior = 0; prior < index; ++prior)
            {
                if (fixture.actions[prior].actionId == action.actionId)
                {
                    return false;
                }
            }
        }
        std::string_view priorSchedule;
        for (const auto& schedule : fixture.schedules)
        {
            if (schedule.scheduleId.empty() || (! priorSchedule.empty() && priorSchedule.compare(schedule.scheduleId) >= 0) ||
                ! IsLowerHex(schedule.scheduleSha256, 64))
            {
                return false;
            }
            priorSchedule = schedule.scheduleId;
            ++scheduleCount;
        }
        if (fixture.schedules.empty() || scheduleCount > kMaximumFixtureRows)
        {
            return false;
        }
        std::string contractSha256;
        if (! ValidateFixtureIdentity(fixture, contractSha256))
        {
            return false;
        }
        if (fixtureIndex != 0)
        {
            manifestCanonical.push_back(',');
        }
        manifestCanonical.append("{\"contractSha256\":");
        AppendJcsString(manifestCanonical, fixture.contractSha256);
        manifestCanonical.append(",\"fixtureId\":");
        AppendJcsString(manifestCanonical, fixture.fixtureId);
        manifestCanonical.push_back('}');
        priorFixture = fixture.fixtureId;
        ++fixtureIndex;
    }
    manifestCanonical.push_back(']');
    std::string manifestSha256;
    constexpr std::array<std::string_view, 5> performanceIds = {
        "alternate-screen-roundtrip",
        "color-style-matrix",
        "synchronized-output",
        "unicode-grapheme-wide",
        "write-response-order",
    };
    const std::vector<uint8_t> performancePayload = BuildPerformancePayload();
    return scheduleCount == 43 && FixtureContract::kPerformancePayload.fixtureIds.size() == performanceIds.size() &&
           std::ranges::equal(FixtureContract::kPerformancePayload.fixtureIds, performanceIds) &&
           performancePayload.size() == FixtureContract::kPerformancePayload.sizeBytes && Sha256Hex(manifestCanonical, manifestSha256) &&
           manifestSha256 == FixtureContract::kFixtureManifestSha256;
}

[[nodiscard]] std::string_view AdapterStringView(RsTerminalGate0String value) noexcept
{
    if (value.ptr == nullptr)
    {
        return value.len == 0 ? std::string_view{} : std::string_view{};
    }
    return std::string_view(reinterpret_cast<const char*>(value.ptr), value.len);
}

[[nodiscard]] bool ValidAdapterString(RsTerminalGate0String value) noexcept
{
    return value.len == 0 || value.ptr != nullptr;
}

[[nodiscard]] bool Utf8ToWide(std::string_view input, std::wstring& output)
{
    if (input.empty() || input.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return false;
    }
    const int required = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, input.data(), static_cast<int>(input.size()), nullptr, 0);
    if (required <= 0)
    {
        return false;
    }
    output.resize(static_cast<size_t>(required));
    return MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, input.data(), static_cast<int>(input.size()), output.data(), required) == required;
}

struct LoadedAdapter
{
    wil::unique_hmodule module;
    RsTerminalGate0AdapterV1 api{};
};

[[nodiscard]] bool LoadAndVerifyAdapter(const Arguments& arguments, LoadedAdapter& loaded)
{
    std::wstring dllPath;
    if (! Utf8ToWide(arguments.adapterDll, dllPath))
    {
        return false;
    }
    loaded.module.reset(LoadLibraryExW(dllPath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_SYSTEM32));
    if (! loaded.module)
    {
        return false;
    }
    const FARPROC address = GetProcAddress(loaded.module.get(), "rs_terminal_gate0_query_adapter");
    if (address == nullptr)
    {
        return false;
    }
    static_assert(sizeof(address) == sizeof(RsTerminalGate0QueryAdapterFn));
    const auto query = std::bit_cast<RsTerminalGate0QueryAdapterFn>(address);

    RsTerminalGate0AdapterV1 wrongVersion{};
    wrongVersion.size          = sizeof(wrongVersion);
    wrongVersion.identity.size = sizeof(wrongVersion.identity);
    if (query(RS_TERMINAL_GATE0_QUERY_SYMBOL_VERSION + 1, RS_TERMINAL_GATE0_ADAPTER_ABI_V1, &wrongVersion) != RS_TERMINAL_GATE0_UNSUPPORTED_ABI)
    {
        return false;
    }
    RsTerminalGate0AdapterV1 wrongAbi{};
    wrongAbi.size          = sizeof(wrongAbi);
    wrongAbi.identity.size = sizeof(wrongAbi.identity);
    if (query(RS_TERMINAL_GATE0_QUERY_SYMBOL_VERSION, RS_TERMINAL_GATE0_ADAPTER_ABI_V1 + 1, &wrongAbi) != RS_TERMINAL_GATE0_UNSUPPORTED_ABI)
    {
        return false;
    }

    loaded.api               = {};
    loaded.api.size          = sizeof(loaded.api);
    loaded.api.identity.size = sizeof(loaded.api.identity);
    if (query(RS_TERMINAL_GATE0_QUERY_SYMBOL_VERSION, arguments.expectedAdapterAbi, &loaded.api) != RS_TERMINAL_GATE0_SUCCESS)
    {
        return false;
    }
    const auto& identity = loaded.api.identity;
    if (loaded.api.size != sizeof(loaded.api) || identity.size != sizeof(identity) || identity.adapter_abi != arguments.expectedAdapterAbi ||
        ! ValidAdapterString(identity.candidate_id) || ! ValidAdapterString(identity.upstream_pin) || ! ValidAdapterString(identity.source_sha256) ||
        ! ValidAdapterString(identity.adapter_header_sha256) || ! ValidAdapterString(identity.build_identity) ||
        AdapterStringView(identity.candidate_id) != arguments.expectedCandidateId ||
        AdapterStringView(identity.upstream_pin) != arguments.expectedUpstreamPin ||
        AdapterStringView(identity.source_sha256) != arguments.expectedSourceSha256 ||
        AdapterStringView(identity.adapter_header_sha256) != arguments.expectedAdapterHeaderSha256 ||
        AdapterStringView(identity.build_identity) != arguments.expectedBuildIdentity || identity.architecture != kExpectedArchitecture ||
        identity.calling_convention != RS_TERMINAL_GATE0_CALL_CDECL || identity.capabilities != arguments.expectedCapabilities ||
        (identity.capabilities & kRequiredCapabilities) != kRequiredCapabilities || identity.pointer_size != sizeof(void*) ||
        loaded.api.session_create == nullptr || loaded.api.session_write == nullptr || loaded.api.session_action == nullptr ||
        loaded.api.session_state_jcs == nullptr || loaded.api.session_destroy == nullptr)
    {
        return false;
    }
    return true;
}

[[nodiscard]] constexpr RsTerminalGate0ResourceLimitsV1 MakeFrozenResourceLimits() noexcept
{
    return {
        .size                               = sizeof(RsTerminalGate0ResourceLimitsV1),
        .terminal_text_session_bytes        = 64U * 1024U * 1024U,
        .terminal_text_process_bytes        = 512U * 1024U * 1024U,
        .image_cpu_session_bytes            = 64U * 1024U * 1024U,
        .image_cpu_process_bytes            = 256U * 1024U * 1024U,
        .image_gpu_session_bytes            = 64U * 1024U * 1024U,
        .image_gpu_process_bytes            = 256U * 1024U * 1024U,
        .pending_effect_session_bytes       = 2U * 1024U * 1024U,
        .pending_effect_process_bytes       = kMaximumSessions * 2U * 1024U * 1024U,
        .pending_effect_session_descriptors = 256,
        .pending_effect_process_descriptors = static_cast<uint32_t>(kMaximumSessions) * 256,
        .maximum_sessions                   = static_cast<uint32_t>(kMaximumSessions),
        .kitty_max_pixels                   = 16'000'000,
        .kitty_max_dimension                = 10'000,
        .reserved                           = 0,
    };
}

class ProcessResourceLedger final
{
public:
    explicit ProcessResourceLedger(const RsTerminalGate0ResourceLimitsV1& limits) noexcept : _limits(limits)
    {
    }

    [[nodiscard]] bool Open(size_t slot, uint64_t& generation) noexcept
    {
        const std::scoped_lock lock(_mutex);
        if (slot >= _sessions.size() || _sessions[slot].active || _activeSessions >= _limits.maximum_sessions)
        {
            return false;
        }
        Session& session = _sessions[slot];
        if (session.generation == std::numeric_limits<uint64_t>::max())
        {
            _arithmeticError = true;
            return false;
        }
        ++session.generation;
        session.active = true;
        ++_activeSessions;
        generation = session.generation;
        return true;
    }

    [[nodiscard]] bool Reserve(size_t slot, uint64_t generation, RsTerminalGate0ResourceKind kind, uint64_t bytes, uint32_t descriptors) noexcept
    {
        const std::scoped_lock lock(_mutex);
        const size_t index = static_cast<size_t>(kind);
        if (! ValidSession(slot, generation) || index >= static_cast<size_t>(RS_TERMINAL_GATE0_RESOURCE_COUNT) ||
            (kind != RS_TERMINAL_GATE0_RESOURCE_PENDING_EFFECT && descriptors != 0))
        {
            ++_rejectedCalls;
            return false;
        }
        Session& session            = _sessions[slot];
        const uint64_t sessionLimit = SessionByteLimit(kind);
        const uint64_t processLimit = ProcessByteLimit(kind);
        if (bytes > sessionLimit - session.currentBytes[index] || bytes > processLimit - _processCurrentBytes[index])
        {
            ++_rejectedCalls;
            return false;
        }
        if (kind == RS_TERMINAL_GATE0_RESOURCE_PENDING_EFFECT)
        {
            if (descriptors > _limits.pending_effect_session_descriptors - session.pendingDescriptors ||
                descriptors > _limits.pending_effect_process_descriptors - _processPendingDescriptors)
            {
                ++_rejectedCalls;
                return false;
            }
            session.pendingDescriptors += descriptors;
            _processPendingDescriptors += descriptors;
            session.peakPendingDescriptors = std::max(session.peakPendingDescriptors, session.pendingDescriptors);
        }
        session.currentBytes[index] += bytes;
        _processCurrentBytes[index] += bytes;
        session.peakBytes[index] = std::max(session.peakBytes[index], session.currentBytes[index]);
        _processPeakBytes[index] = std::max(_processPeakBytes[index], _processCurrentBytes[index]);
        ++_reserveCalls;
        return true;
    }

    [[nodiscard]] bool Release(size_t slot, uint64_t generation, RsTerminalGate0ResourceKind kind, uint64_t bytes, uint32_t descriptors) noexcept
    {
        const std::scoped_lock lock(_mutex);
        const size_t index = static_cast<size_t>(kind);
        if (! ValidSession(slot, generation) || index >= static_cast<size_t>(RS_TERMINAL_GATE0_RESOURCE_COUNT) ||
            (kind != RS_TERMINAL_GATE0_RESOURCE_PENDING_EFFECT && descriptors != 0))
        {
            _arithmeticError = true;
            return false;
        }
        Session& session = _sessions[slot];
        if (bytes > session.currentBytes[index] || bytes > _processCurrentBytes[index] || descriptors > session.pendingDescriptors ||
            descriptors > _processPendingDescriptors)
        {
            _arithmeticError = true;
            return false;
        }
        session.currentBytes[index] -= bytes;
        _processCurrentBytes[index] -= bytes;
        if (kind == RS_TERMINAL_GATE0_RESOURCE_PENDING_EFFECT)
        {
            session.pendingDescriptors -= descriptors;
            _processPendingDescriptors -= descriptors;
        }
        ++_releaseCalls;
        return true;
    }

    [[nodiscard]] bool Close(size_t slot, uint64_t generation) noexcept
    {
        const std::scoped_lock lock(_mutex);
        if (! ValidSession(slot, generation))
        {
            return false;
        }
        Session& session = _sessions[slot];
        if (std::ranges::any_of(session.currentBytes, [](uint64_t value) { return value != 0; }) || session.pendingDescriptors != 0)
        {
            return false;
        }
        session.active = false;
        --_activeSessions;
        return true;
    }

    [[nodiscard]] bool CheckedArithmetic() const noexcept
    {
        const std::scoped_lock lock(_mutex);
        return ! _arithmeticError;
    }

    [[nodiscard]] bool AllReleased() const noexcept
    {
        const std::scoped_lock lock(_mutex);
        return _activeSessions == 0 && _processPendingDescriptors == 0 && std::ranges::all_of(_processCurrentBytes, [](uint64_t value) { return value == 0; });
    }

    [[nodiscard]] bool PeaksWithinLimits() const noexcept
    {
        const std::scoped_lock lock(_mutex);
        for (size_t index = 0; index < _processPeakBytes.size(); ++index)
        {
            const auto kind = static_cast<RsTerminalGate0ResourceKind>(index);
            if (_processPeakBytes[index] > ProcessByteLimit(kind))
            {
                return false;
            }
        }
        return true;
    }

    [[nodiscard]] uint64_t CurrentCpuBytes() const noexcept
    {
        const std::scoped_lock lock(_mutex);
        uint64_t total                                                = 0;
        constexpr std::array<RsTerminalGate0ResourceKind, 3> cpuKinds = {
            RS_TERMINAL_GATE0_RESOURCE_TERMINAL_TEXT_CPU,
            RS_TERMINAL_GATE0_RESOURCE_IMAGE_CPU,
            RS_TERMINAL_GATE0_RESOURCE_PENDING_EFFECT,
        };
        for (const auto kind : cpuKinds)
        {
            const uint64_t current = _processCurrentBytes[static_cast<size_t>(kind)];
            if (current > std::numeric_limits<uint64_t>::max() - total)
            {
                return std::numeric_limits<uint64_t>::max();
            }
            total += current;
        }
        return total;
    }

private:
    struct Session
    {
        std::array<uint64_t, RS_TERMINAL_GATE0_RESOURCE_COUNT> currentBytes{};
        std::array<uint64_t, RS_TERMINAL_GATE0_RESOURCE_COUNT> peakBytes{};
        uint64_t generation             = 0;
        uint32_t pendingDescriptors     = 0;
        uint32_t peakPendingDescriptors = 0;
        bool active                     = false;
    };

    [[nodiscard]] bool ValidSession(size_t slot, uint64_t generation) const noexcept
    {
        return slot < _sessions.size() && _sessions[slot].active && _sessions[slot].generation == generation;
    }

    [[nodiscard]] uint64_t SessionByteLimit(RsTerminalGate0ResourceKind kind) const noexcept
    {
        switch (kind)
        {
            case RS_TERMINAL_GATE0_RESOURCE_TERMINAL_TEXT_CPU: return _limits.terminal_text_session_bytes;
            case RS_TERMINAL_GATE0_RESOURCE_IMAGE_CPU: return _limits.image_cpu_session_bytes;
            case RS_TERMINAL_GATE0_RESOURCE_IMAGE_GPU: return _limits.image_gpu_session_bytes;
            case RS_TERMINAL_GATE0_RESOURCE_PENDING_EFFECT: return _limits.pending_effect_session_bytes;
            case RS_TERMINAL_GATE0_RESOURCE_COUNT: break;
        }
        return 0;
    }

    [[nodiscard]] uint64_t ProcessByteLimit(RsTerminalGate0ResourceKind kind) const noexcept
    {
        switch (kind)
        {
            case RS_TERMINAL_GATE0_RESOURCE_TERMINAL_TEXT_CPU: return _limits.terminal_text_process_bytes;
            case RS_TERMINAL_GATE0_RESOURCE_IMAGE_CPU: return _limits.image_cpu_process_bytes;
            case RS_TERMINAL_GATE0_RESOURCE_IMAGE_GPU: return _limits.image_gpu_process_bytes;
            case RS_TERMINAL_GATE0_RESOURCE_PENDING_EFFECT: return _limits.pending_effect_process_bytes;
            case RS_TERMINAL_GATE0_RESOURCE_COUNT: break;
        }
        return 0;
    }

    const RsTerminalGate0ResourceLimitsV1& _limits;
    mutable std::mutex _mutex;
    std::array<Session, kMaximumSessions> _sessions{};
    std::array<uint64_t, RS_TERMINAL_GATE0_RESOURCE_COUNT> _processCurrentBytes{};
    std::array<uint64_t, RS_TERMINAL_GATE0_RESOURCE_COUNT> _processPeakBytes{};
    uint32_t _processPendingDescriptors = 0;
    uint32_t _activeSessions            = 0;
    uint64_t _reserveCalls              = 0;
    uint64_t _releaseCalls              = 0;
    uint64_t _rejectedCalls             = 0;
    bool _arithmeticError               = false;
};

struct ResourceCallbackContext
{
    ProcessResourceLedger* ledger   = nullptr;
    size_t slot                     = 0;
    uint64_t generation             = 0;
    std::atomic<bool> accepting     = false;
    std::atomic<bool> callbackError = false;
};

bool RS_TERMINAL_GATE0_CALL ReserveResource(void* userdata, RsTerminalGate0ResourceKind kind, uint64_t bytes, uint32_t descriptors) noexcept
{
    auto& context = *static_cast<ResourceCallbackContext*>(userdata);
    if (! context.accepting.load(std::memory_order_acquire) || context.ledger == nullptr)
    {
        context.callbackError.store(true, std::memory_order_release);
        return false;
    }
    return context.ledger->Reserve(context.slot, context.generation, kind, bytes, descriptors);
}

void RS_TERMINAL_GATE0_CALL ReleaseResource(void* userdata, RsTerminalGate0ResourceKind kind, uint64_t bytes, uint32_t descriptors) noexcept
{
    auto& context = *static_cast<ResourceCallbackContext*>(userdata);
    if (! context.accepting.load(std::memory_order_acquire) || context.ledger == nullptr ||
        ! context.ledger->Release(context.slot, context.generation, kind, bytes, descriptors))
    {
        context.callbackError.store(true, std::memory_order_release);
    }
}

[[nodiscard]] RsTerminalGate0String AdapterString(std::string_view value) noexcept
{
    return {
        .ptr = reinterpret_cast<const uint8_t*>(value.data()),
        .len = value.size(),
    };
}

struct FeedAudit
{
    uint64_t bytes = 0;
    uint64_t calls = 0;
};

[[nodiscard]] bool WriteChunk(const RsTerminalGate0AdapterV1& adapter,
                              RsTerminalGate0Session session,
                              std::span<const uint8_t> payload,
                              size_t begin,
                              size_t end,
                              FeedAudit& audit) noexcept
{
    if (begin > end || end > payload.size() || adapter.session_write(session, payload.data() + begin, end - begin) != RS_TERMINAL_GATE0_SUCCESS)
    {
        return false;
    }
    audit.bytes += static_cast<uint64_t>(end - begin);
    ++audit.calls;
    return true;
}

[[nodiscard]] bool FeedSchedule(const RsTerminalGate0AdapterV1& adapter,
                                RsTerminalGate0Session session,
                                std::span<const uint8_t> payload,
                                const FixtureContract::ScheduleSpec& schedule,
                                FeedAudit& audit) noexcept
{
    audit = {};
    switch (schedule.algorithm)
    {
        case FixtureContract::ScheduleAlgorithm::AllAtOnce:
            if (! WriteChunk(adapter, session, payload, 0, payload.size(), audit))
            {
                return false;
            }
            break;
        case FixtureContract::ScheduleAlgorithm::Bytewise:
            for (size_t offset = 0; offset < payload.size(); ++offset)
            {
                if (! WriteChunk(adapter, session, payload, offset, offset + 1, audit))
                {
                    return false;
                }
            }
            break;
        case FixtureContract::ScheduleAlgorithm::Fibonacci:
        {
            constexpr std::array<size_t, 10> chunks = {
                1,
                2,
                3,
                5,
                8,
                13,
                21,
                34,
                55,
                89,
            };
            size_t offset     = 0;
            size_t chunkIndex = 0;
            while (offset < payload.size())
            {
                const size_t end = offset + std::min(chunks[chunkIndex], payload.size() - offset);
                if (! WriteChunk(adapter, session, payload, offset, end, audit))
                {
                    return false;
                }
                offset     = end;
                chunkIndex = (chunkIndex + 1) % chunks.size();
            }
            break;
        }
        case FixtureContract::ScheduleAlgorithm::TerminatorCuts:
        {
            size_t prior = 0;
            for (const uint64_t cut : schedule.cutOffsets)
            {
                if (cut > payload.size() || cut <= prior || ! WriteChunk(adapter, session, payload, prior, static_cast<size_t>(cut), audit))
                {
                    return false;
                }
                prior = static_cast<size_t>(cut);
            }
            if (! WriteChunk(adapter, session, payload, prior, payload.size(), audit))
            {
                return false;
            }
            break;
        }
        default: return false;
    }
    return audit.calls != 0 && audit.bytes == payload.size();
}

[[nodiscard]] bool IsLimitFixture(std::string_view fixtureId) noexcept
{
    constexpr std::array<std::string_view, 11> ids = {
        "hyperlink-boundaries",
        "kitty-dimension-overflow",
        "kitty-pixel-boundary",
        "kitty-png-bomb",
        "kitty-replacement-transient",
        "kitty-zlib-bomb",
        "nonkitty-generic-boundary",
        "osc52-boundaries",
        "pending-effects-boundary",
        "scrollback-byte-boundary",
        "title-cwd-boundaries",
    };
    return std::ranges::find(ids, fixtureId) != ids.end();
}

[[nodiscard]] bool IsLimitAction(std::string_view actionId) noexcept
{
    constexpr std::array<std::string_view, 8> ids = {
        "probe-osc8-limits",
        "probe-kitty-pixel-limit",
        "probe-kitty-replacement-ledger",
        "probe-generic-escape-limit",
        "probe-osc52-limits",
        "probe-pending-effect-ledger",
        "probe-terminal-text-ledger",
        "probe-title-cwd-limits",
    };
    return std::ranges::find(ids, actionId) != ids.end();
}

[[nodiscard]] bool IsOrderingAction(std::string_view actionId) noexcept
{
    return actionId == "encode-paste-enabled" || actionId == "encode-focus-gained-lost" || actionId == "encode-mouse-matrix" ||
           actionId == "admit-external-events";
}

[[nodiscard]] bool IsKnownNonFeedAction(std::string_view actionId) noexcept
{
    constexpr std::array<std::string_view, 19> ids = {
        "admit-external-events",       "encode-focus-gained-lost",
        "encode-mouse-matrix",         "encode-paste-disabled",
        "encode-paste-enabled",        "probe-generic-escape-limit",
        "probe-kitty-pixel-limit",     "probe-kitty-replacement-ledger",
        "probe-osc52-limits",          "probe-osc8-limits",
        "probe-pending-effect-ledger", "probe-terminal-text-ledger",
        "probe-title-cwd-limits",      "resize-sequence",
        "select-logical-range",        "snapshot-copy-fail-retry",
        "snapshot-end-mutate-probe",   "snapshot-static-kitty",
        "teardown-sequence",
    };
    return std::ranges::find(ids, actionId) != ids.end();
}

[[nodiscard]] bool ValidateActionObservation(std::string_view actionId, const RsTerminalGate0ActionObservationV1& observation, bool& limitEvidence) noexcept
{
    constexpr uint64_t knownFlags = RS_TERMINAL_GATE0_ACTION_OBSERVED | RS_TERMINAL_GATE0_ACTION_CHECKED_ARITHMETIC |
                                    RS_TERMINAL_GATE0_ACTION_BOUNDARY_ACCEPTED | RS_TERMINAL_GATE0_ACTION_PLUS_ONE_REJECTED |
                                    RS_TERMINAL_GATE0_ACTION_REJECTION_BEFORE_ALLOCATION | RS_TERMINAL_GATE0_ACTION_NO_PARTIAL_EFFECT |
                                    RS_TERMINAL_GATE0_ACTION_ORDER_PRESERVED | RS_TERMINAL_GATE0_ACTION_QUIET_POINT_REACHED;
    if (observation.size != sizeof(observation) || (observation.flags & ~knownFlags) != 0 ||
        (observation.flags & (RS_TERMINAL_GATE0_ACTION_OBSERVED | RS_TERMINAL_GATE0_ACTION_CHECKED_ARITHMETIC)) !=
            (RS_TERMINAL_GATE0_ACTION_OBSERVED | RS_TERMINAL_GATE0_ACTION_CHECKED_ARITHMETIC))
    {
        return false;
    }
    if (IsLimitAction(actionId))
    {
        constexpr uint64_t required = RS_TERMINAL_GATE0_ACTION_BOUNDARY_ACCEPTED | RS_TERMINAL_GATE0_ACTION_PLUS_ONE_REJECTED |
                                      RS_TERMINAL_GATE0_ACTION_REJECTION_BEFORE_ALLOCATION | RS_TERMINAL_GATE0_ACTION_NO_PARTIAL_EFFECT;
        if ((observation.flags & required) != required || observation.reservation_calls_before_rejection != 0 ||
            observation.allocation_attempts_before_rejection != 0)
        {
            return false;
        }
        limitEvidence = true;
    }
    if (IsOrderingAction(actionId) && (observation.flags & RS_TERMINAL_GATE0_ACTION_ORDER_PRESERVED) == 0)
    {
        return false;
    }
    if (actionId == "teardown-sequence" && (observation.flags & RS_TERMINAL_GATE0_ACTION_QUIET_POINT_REACHED) == 0)
    {
        return false;
    }
    return true;
}

[[nodiscard]] bool ReadStateJcs(const RsTerminalGate0AdapterV1& adapter, RsTerminalGate0Session session, std::string& state)
{
    size_t required = 0;
    if (adapter.session_state_jcs(session, nullptr, 0, &required) != RS_TERMINAL_GATE0_OUT_OF_SPACE || required == 0 || required > 1024U * 1024U)
    {
        return false;
    }
    state.assign(required, '\0');
    size_t written = 0;
    if (adapter.session_state_jcs(session, reinterpret_cast<uint8_t*>(state.data()), state.size(), &written) != RS_TERMINAL_GATE0_SUCCESS ||
        written != required)
    {
        return false;
    }
    return true;
}

[[nodiscard]] bool ResourceTraceJcs(
    bool allReservationsReleased, bool checkedArithmetic, bool limitFixture, bool peakWithinFrozenLimits, bool rejectionBeforeAllocation, std::string& output)
{
    output.clear();
    output.append("{\"allReservationsReleased\":");
    output.append(allReservationsReleased ? "true" : "false");
    output.append(",\"checkedArithmetic\":");
    output.append(checkedArithmetic ? "true" : "false");
    output.append(",\"limitFixture\":");
    output.append(limitFixture ? "true" : "false");
    output.append(",\"peakWithinFrozenLimits\":");
    output.append(peakWithinFrozenLimits ? "true" : "false");
    output.append(",\"rejectionBeforeAllocation\":");
    output.append(rejectionBeforeAllocation ? "true" : "false");
    output.push_back('}');
    return true;
}

struct FixtureRow
{
    std::string_view fixtureId;
    std::string_view scheduleId;
    std::string_view inputSha256;
    uint64_t inputSizeBytes = 0;
    std::string_view expectedStateSha256;
    std::string actualStateSha256;
    std::string_view expectedResourceTraceSha256;
    std::string actualResourceTraceSha256;
    bool passed                = false;
    const char* resultCategory = "";
};

[[nodiscard]] bool RunFixtureRow(const RsTerminalGate0AdapterV1& adapter,
                                 const RsTerminalGate0ResourceLimitsV1& limits,
                                 ProcessResourceLedger& ledger,
                                 ResourceCallbackContext& callbackContext,
                                 const FixtureContract::Fixture& fixture,
                                 const FixtureContract::ScheduleSpec& schedule,
                                 std::span<const uint8_t> payload,
                                 FixtureRow& row)
{
    row.fixtureId                = fixture.fixtureId;
    row.scheduleId               = schedule.scheduleId;
    row.inputSha256              = fixture.inputSha256;
    row.inputSizeBytes           = fixture.inputSizeBytes;
    const auto* expectedState    = FindCheck(fixture, "state");
    const auto* expectedResource = FindCheck(fixture, "resource-trace");
    if (expectedState == nullptr || expectedResource == nullptr)
    {
        row.resultCategory = "fixture-check-set-invalid";
        return false;
    }
    row.expectedStateSha256         = expectedState->expectedSha256;
    row.expectedResourceTraceSha256 = expectedResource->expectedSha256;

    uint64_t generation = 0;
    if (! ledger.Open(0, generation))
    {
        row.resultCategory = "resource-session-open-failed";
        return false;
    }
    callbackContext.ledger     = &ledger;
    callbackContext.slot       = 0;
    callbackContext.generation = generation;
    callbackContext.callbackError.store(false, std::memory_order_relaxed);
    callbackContext.accepting.store(true, std::memory_order_release);
    const RsTerminalGate0ResourceCallbacksV1 callbacks = {
        .size     = sizeof(RsTerminalGate0ResourceCallbacksV1),
        .userdata = &callbackContext,
        .reserve  = ReserveResource,
        .release  = ReleaseResource,
    };
    const RsTerminalGate0SessionConfigV1 config = {
        .size               = sizeof(RsTerminalGate0SessionConfigV1),
        .columns            = kColumns,
        .rows               = kRows,
        .cell_width_px      = 8,
        .cell_height_px     = 16,
        .fixture_id         = AdapterString(fixture.fixtureId),
        .schedule_id        = AdapterString(schedule.scheduleId),
        .limits             = &limits,
        .resource_callbacks = &callbacks,
    };

    bool createSucceeded           = false;
    bool actionsPassed             = true;
    bool feedPassed                = false;
    bool stateRead                 = false;
    bool stateExact                = false;
    bool stateHashValid            = false;
    bool limitEvidence             = false;
    bool destroySucceeded          = false;
    RsTerminalGate0Session session = nullptr;
    if (adapter.session_create(&config, &session) == RS_TERMINAL_GATE0_SUCCESS && session != nullptr)
    {
        createSucceeded = true;
    }
    else
    {
        actionsPassed = false;
    }

    if (createSucceeded)
    {
        for (size_t index = 0; index < fixture.actions.size(); ++index)
        {
            const auto& frozenAction = fixture.actions[index];
            if (index == 0)
            {
                FeedAudit audit{};
                feedPassed    = frozenAction.actionId == "feed-current-schedule" && FeedSchedule(adapter, session, payload, schedule, audit);
                actionsPassed = actionsPassed && feedPassed;
                continue;
            }
            if (! IsKnownNonFeedAction(frozenAction.actionId))
            {
                actionsPassed = false;
                continue;
            }
            const RsTerminalGate0ActionV1 action = {
                .size          = sizeof(RsTerminalGate0ActionV1),
                .action_id     = AdapterString(frozenAction.actionId),
                .arguments_jcs = AdapterString(frozenAction.argumentsJcs),
            };
            RsTerminalGate0ActionObservationV1 observation{};
            observation.size           = sizeof(observation);
            const bool actionSucceeded = adapter.session_action(session, &action, &observation) == RS_TERMINAL_GATE0_SUCCESS &&
                                         ValidateActionObservation(frozenAction.actionId, observation, limitEvidence);
            actionsPassed              = actionsPassed && actionSucceeded;
        }

        std::string actualState;
        stateRead = ReadStateJcs(adapter, session, actualState);
        if (stateRead)
        {
            stateHashValid = Sha256Hex(actualState, row.actualStateSha256);
            stateExact     = actualState == expectedState->expectedValue && stateHashValid && row.actualStateSha256 == expectedState->expectedSha256;
        }
        if (! stateHashValid)
        {
            (void)Sha256Hex("{\"stateUnavailable\":true}", row.actualStateSha256);
        }
    }
    if (session != nullptr)
    {
        destroySucceeded = adapter.session_destroy(session) == RS_TERMINAL_GATE0_SUCCESS;
    }

    callbackContext.accepting.store(false, std::memory_order_release);
    const bool sessionReleased           = ledger.Close(0, generation);
    const bool callbackClean             = ! callbackContext.callbackError.load(std::memory_order_acquire);
    const bool allReleased               = destroySucceeded && sessionReleased && callbackClean && ledger.AllReleased();
    const bool checkedArithmetic         = actionsPassed && ledger.CheckedArithmetic() && callbackClean;
    const bool limitFixture              = IsLimitFixture(fixture.fixtureId);
    const bool rejectionBeforeAllocation = limitFixture && stateExact && (limitEvidence || fixture.actions.size() == 1);
    std::string actualResource;
    const bool resourceBuilt =
        ResourceTraceJcs(allReleased, checkedArithmetic, limitFixture, ledger.PeaksWithinLimits(), rejectionBeforeAllocation, actualResource);
    const bool resourceHashValid = resourceBuilt && Sha256Hex(actualResource, row.actualResourceTraceSha256);
    const bool resourceExact =
        resourceHashValid && actualResource == expectedResource->expectedValue && row.actualResourceTraceSha256 == expectedResource->expectedSha256;
    row.passed = createSucceeded && feedPassed && actionsPassed && stateExact && resourceExact && allReleased;
    if (! createSucceeded)
    {
        row.resultCategory = "session-create-failed";
    }
    else if (! feedPassed)
    {
        row.resultCategory = "scheduled-feed-failed";
    }
    else if (! actionsPassed)
    {
        row.resultCategory = "declared-action-failed";
    }
    else if (! stateRead)
    {
        row.resultCategory = "state-copy-failed";
    }
    else if (! stateExact)
    {
        row.resultCategory = "state-oracle-mismatch";
    }
    else if (! allReleased)
    {
        row.resultCategory = "session-not-quiet-or-balanced";
    }
    else if (! resourceExact)
    {
        row.resultCategory = "resource-trace-mismatch";
    }
    else
    {
        row.resultCategory = "state-and-resource-trace-matched";
    }
    return row.passed;
}

[[nodiscard]] bool RunFixtureMatrix(const RsTerminalGate0AdapterV1& adapter,
                                    const RsTerminalGate0ResourceLimitsV1& limits,
                                    ProcessResourceLedger& ledger,
                                    std::span<ResourceCallbackContext> callbackContexts,
                                    std::vector<FixtureRow>& rows)
{
    rows.clear();
    rows.reserve(43);
    size_t contextIndex = 0;
    bool passed         = true;
    for (const auto& fixture : FixtureContract::GetFixtures())
    {
        std::vector<uint8_t> payload;
        if (! ExpandFixtureInput(fixture, payload))
        {
            return false;
        }
        for (const auto& schedule : fixture.schedules)
        {
            if (contextIndex >= callbackContexts.size())
            {
                return false;
            }
            FixtureRow row;
            const bool rowPassed = RunFixtureRow(adapter, limits, ledger, callbackContexts[contextIndex], fixture, schedule, payload, row);
            rows.push_back(std::move(row));
            passed = passed && rowPassed;
            ++contextIndex;
        }
    }
    if (rows.size() != 43)
    {
        return false;
    }
    for (size_t index = 0; index < rows.size(); ++index)
    {
        if (callbackContexts[index].callbackError.load(std::memory_order_acquire))
        {
            rows[index].passed         = false;
            rows[index].resultCategory = "callback-after-destroy";
            passed                     = false;
        }
    }
    return passed;
}

[[nodiscard]] bool RunEightSessionStorm(const RsTerminalGate0AdapterV1& adapter,
                                        const RsTerminalGate0ResourceLimitsV1& limits,
                                        ProcessResourceLedger& ledger,
                                        std::span<ResourceCallbackContext> callbackContexts) noexcept
{
    constexpr size_t count = 8;
    if (callbackContexts.size() < count)
    {
        return false;
    }
    std::array<RsTerminalGate0ResourceCallbacksV1, count> callbacks{};
    std::array<RsTerminalGate0SessionConfigV1, count> configs{};
    std::array<RsTerminalGate0Session, count> sessions{};
    std::array<uint64_t, count> generations{};
    std::array<bool, count> opened{};
    bool passed                           = true;
    constexpr std::string_view fixtureId  = "teardown-callback-quiet";
    constexpr std::string_view scheduleId = "eight-session-storm";
    for (size_t index = 0; index < count; ++index)
    {
        auto& context = callbackContexts[index];
        opened[index] = ledger.Open(index, generations[index]);
        if (! opened[index])
        {
            passed = false;
            break;
        }
        context.ledger     = &ledger;
        context.slot       = index;
        context.generation = generations[index];
        context.callbackError.store(false, std::memory_order_relaxed);
        context.accepting.store(true, std::memory_order_release);
        callbacks[index] = {
            .size     = sizeof(RsTerminalGate0ResourceCallbacksV1),
            .userdata = &context,
            .reserve  = ReserveResource,
            .release  = ReleaseResource,
        };
        configs[index] = {
            .size               = sizeof(RsTerminalGate0SessionConfigV1),
            .columns            = kColumns,
            .rows               = kRows,
            .cell_width_px      = 8,
            .cell_height_px     = 16,
            .fixture_id         = AdapterString(fixtureId),
            .schedule_id        = AdapterString(scheduleId),
            .limits             = &limits,
            .resource_callbacks = &callbacks[index],
        };
        if (adapter.session_create(&configs[index], &sessions[index]) != RS_TERMINAL_GATE0_SUCCESS || sessions[index] == nullptr)
        {
            passed = false;
            break;
        }
        constexpr std::array<uint8_t, 5> storm = {
            's',
            't',
            'o',
            'r',
            'm',
        };
        if (adapter.session_write(sessions[index], storm.data(), storm.size()) != RS_TERMINAL_GATE0_SUCCESS)
        {
            passed = false;
            break;
        }
    }
    for (size_t reverse = count; reverse != 0; --reverse)
    {
        const size_t index = reverse - 1;
        if (sessions[index] != nullptr && adapter.session_destroy(sessions[index]) != RS_TERMINAL_GATE0_SUCCESS)
        {
            passed = false;
        }
        callbackContexts[index].accepting.store(false, std::memory_order_release);
        if (opened[index] && ! ledger.Close(index, generations[index]))
        {
            passed = false;
        }
    }
    for (size_t index = 0; index < count; ++index)
    {
        if (callbackContexts[index].callbackError.load(std::memory_order_acquire))
        {
            passed = false;
        }
    }
    return passed && ledger.CheckedArithmetic() && ledger.PeaksWithinLimits() && ledger.AllReleased();
}

[[nodiscard]] const FixtureContract::Fixture* FindFixture(std::string_view fixtureId) noexcept
{
    const auto fixtures = FixtureContract::GetFixtures();
    const auto found    = std::ranges::find(fixtures, fixtureId, &FixtureContract::Fixture::fixtureId);
    return found == fixtures.end() ? nullptr : &*found;
}

[[nodiscard]] std::vector<uint8_t> BuildPerformancePayload()
{
    const auto& specification = FixtureContract::kPerformancePayload;
    std::array<std::vector<uint8_t>, 5> fixturePayloads;
    if (specification.fixtureIds.size() != fixturePayloads.size())
    {
        return {};
    }
    for (size_t index = 0; index < fixturePayloads.size(); ++index)
    {
        const auto* fixture = FindFixture(specification.fixtureIds[index]);
        if (fixture == nullptr || ! ExpandFixtureInput(*fixture, fixturePayloads[index]))
        {
            return {};
        }
    }
    if (specification.sizeBytes > std::numeric_limits<size_t>::max())
    {
        return {};
    }
    std::vector<uint8_t> payload;
    payload.reserve(static_cast<size_t>(specification.sizeBytes));
    uint64_t counter = 0;
    while (payload.size() < static_cast<size_t>(specification.sizeBytes))
    {
        const auto& fixture = fixturePayloads[static_cast<size_t>(counter % fixturePayloads.size())];
        std::array<char, 16> counterBytes{};
        std::array<char, 16> convertedBytes{};
        const auto converted = std::to_chars(convertedBytes.data(), convertedBytes.data() + convertedBytes.size(), counter, 16);
        if (converted.ec != std::errc{})
        {
            return {};
        }
        const size_t digits = static_cast<size_t>(converted.ptr - convertedBytes.data());
        std::fill(counterBytes.begin(), counterBytes.end() - static_cast<ptrdiff_t>(digits), '0');
        std::copy(convertedBytes.begin(), convertedBytes.begin() + static_cast<ptrdiff_t>(digits), counterBytes.end() - static_cast<ptrdiff_t>(digits));
        constexpr size_t suffixLength = 18;
        const size_t remaining        = static_cast<size_t>(specification.sizeBytes) - payload.size();
        if (fixture.size() + suffixLength > remaining)
        {
            payload.insert(payload.end(), remaining, 'x');
            break;
        }
        payload.insert(payload.end(), fixture.begin(), fixture.end());
        payload.push_back('#');
        payload.insert(
            payload.end(), reinterpret_cast<const uint8_t*>(counterBytes.data()), reinterpret_cast<const uint8_t*>(counterBytes.data() + counterBytes.size()));
        payload.push_back('\n');
        ++counter;
    }
    std::string sha256;
    if (payload.size() != specification.sizeBytes || ! Sha256Hex(payload, sha256) || sha256 != specification.sha256)
    {
        return {};
    }
    return payload;
}

struct PerformanceSession
{
    RsTerminalGate0Session session   = nullptr;
    ResourceCallbackContext* context = nullptr;
    RsTerminalGate0ResourceCallbacksV1 callbacks{};
    uint64_t generation = 0;
    size_t slot         = 0;
};

[[nodiscard]] bool CreatePerformanceSession(const RsTerminalGate0AdapterV1& adapter,
                                            const RsTerminalGate0ResourceLimitsV1& limits,
                                            ProcessResourceLedger& ledger,
                                            ResourceCallbackContext& context,
                                            size_t slot,
                                            std::string_view fixtureId,
                                            std::string_view scheduleId,
                                            PerformanceSession& session) noexcept
{
    session         = {};
    session.context = &context;
    session.slot    = slot;
    if (! ledger.Open(slot, session.generation))
    {
        return false;
    }
    context.ledger     = &ledger;
    context.slot       = slot;
    context.generation = session.generation;
    context.callbackError.store(false, std::memory_order_relaxed);
    context.accepting.store(true, std::memory_order_release);
    session.callbacks = {
        .size     = sizeof(RsTerminalGate0ResourceCallbacksV1),
        .userdata = &context,
        .reserve  = ReserveResource,
        .release  = ReleaseResource,
    };
    const RsTerminalGate0SessionConfigV1 config = {
        .size               = sizeof(RsTerminalGate0SessionConfigV1),
        .columns            = kColumns,
        .rows               = kRows,
        .cell_width_px      = 8,
        .cell_height_px     = 16,
        .fixture_id         = AdapterString(fixtureId),
        .schedule_id        = AdapterString(scheduleId),
        .limits             = &limits,
        .resource_callbacks = &session.callbacks,
    };
    if (adapter.session_create(&config, &session.session) != RS_TERMINAL_GATE0_SUCCESS || session.session == nullptr)
    {
        context.accepting.store(false, std::memory_order_release);
        (void)ledger.Close(slot, session.generation);
        return false;
    }
    return true;
}

[[nodiscard]] bool DestroyPerformanceSession(const RsTerminalGate0AdapterV1& adapter, ProcessResourceLedger& ledger, PerformanceSession& session) noexcept
{
    if (session.context == nullptr)
    {
        return false;
    }
    const bool destroyed = session.session != nullptr && adapter.session_destroy(session.session) == RS_TERMINAL_GATE0_SUCCESS;
    session.context->accepting.store(false, std::memory_order_release);
    const bool closed        = ledger.Close(session.slot, session.generation);
    const bool callbackClean = ! session.context->callbackError.load(std::memory_order_acquire);
    session.session          = nullptr;
    return destroyed && closed && callbackClean;
}

[[nodiscard]] bool RunBenchmarkAction(const RsTerminalGate0AdapterV1& adapter,
                                      RsTerminalGate0Session session,
                                      std::string_view actionId,
                                      std::string_view argumentsJcs) noexcept
{
    const RsTerminalGate0ActionV1 action = {
        .size          = sizeof(RsTerminalGate0ActionV1),
        .action_id     = AdapterString(actionId),
        .arguments_jcs = AdapterString(argumentsJcs),
    };
    RsTerminalGate0ActionObservationV1 observation{};
    observation.size            = sizeof(observation);
    constexpr uint64_t required = RS_TERMINAL_GATE0_ACTION_OBSERVED | RS_TERMINAL_GATE0_ACTION_CHECKED_ARITHMETIC;
    return adapter.session_action(session, &action, &observation) == RS_TERMINAL_GATE0_SUCCESS && observation.size == sizeof(observation) &&
           (observation.flags & required) == required;
}

struct Measurement
{
    const char* metricId = "";
    const char* unit     = "";
    uint32_t warmupCount = 0;
    std::array<uint64_t, 50> samples{};
    size_t sampleCount = 0;
};

[[nodiscard]] uint64_t QpcDeltaToMicroseconds(uint64_t delta, uint64_t frequency) noexcept
{
    if (frequency == 0)
    {
        return std::numeric_limits<uint64_t>::max();
    }
    const uint64_t whole     = delta / frequency;
    const uint64_t remainder = delta % frequency;
    if (whole > std::numeric_limits<uint64_t>::max() / 1'000'000U || remainder > std::numeric_limits<uint64_t>::max() / 1'000'000U)
    {
        return std::numeric_limits<uint64_t>::max();
    }
    const uint64_t base                = whole * 1'000'000U;
    const uint64_t numerator           = remainder * 1'000'000U;
    uint64_t fractional                = numerator / frequency;
    const uint64_t fractionalRemainder = numerator % frequency;
    const bool greaterThanHalf         = fractionalRemainder > frequency - fractionalRemainder;
    const bool exactHalf               = fractionalRemainder == frequency - fractionalRemainder;
    if (greaterThanHalf || (exactHalf && (fractional & 1U) != 0))
    {
        ++fractional;
    }
    if (fractional > std::numeric_limits<uint64_t>::max() - base)
    {
        return std::numeric_limits<uint64_t>::max();
    }
    return base + fractional;
}

template <typename Prepare, typename Scenario, typename Cleanup>
[[nodiscard]] bool CollectMeasurement(
    Measurement& measurement, uint32_t warmups, size_t samples, uint64_t frequency, Prepare&& prepare, Scenario&& scenario, Cleanup&& cleanup)
{
    if (samples > measurement.samples.size())
    {
        return false;
    }
    measurement.warmupCount = warmups;
    measurement.sampleCount = samples;
    const size_t iterations = static_cast<size_t>(warmups) + samples;
    for (size_t iteration = 0; iteration < iterations; ++iteration)
    {
        if (! prepare())
        {
            (void)cleanup();
            return false;
        }
        LARGE_INTEGER begin{};
        LARGE_INTEGER end{};
        const bool began          = QueryPerformanceCounter(&begin) != FALSE;
        const bool scenarioPassed = began && scenario();
        const bool ended          = QueryPerformanceCounter(&end) != FALSE;
        const bool cleanupPassed  = cleanup();
        if (! scenarioPassed || ! ended || ! cleanupPassed || end.QuadPart < begin.QuadPart)
        {
            return false;
        }
        if (iteration >= warmups)
        {
            const uint64_t converted = QpcDeltaToMicroseconds(static_cast<uint64_t>(end.QuadPart - begin.QuadPart), frequency);
            if (converted == std::numeric_limits<uint64_t>::max())
            {
                return false;
            }
            measurement.samples[iteration - warmups] = converted;
        }
    }
    return true;
}

[[nodiscard]] bool RunFullMeasurements(const Arguments& arguments,
                                       const RsTerminalGate0AdapterV1& adapter,
                                       const RsTerminalGate0ResourceLimitsV1& limits,
                                       ProcessResourceLedger& ledger,
                                       std::span<ResourceCallbackContext> contexts,
                                       std::array<Measurement, 11>& measurements,
                                       uint64_t& qpcFrequency)
{
    LARGE_INTEGER frequency{};
    if (QueryPerformanceFrequency(&frequency) == FALSE || frequency.QuadPart <= 0)
    {
        return false;
    }
    qpcFrequency                            = static_cast<uint64_t>(frequency.QuadPart);
    const std::vector<uint8_t> parsePayload = BuildPerformancePayload();
    if (parsePayload.empty())
    {
        return false;
    }
    std::vector<uint8_t> snapshotPayload;
    std::vector<uint8_t> resizePayload;
    auto appendBytes = [](std::vector<uint8_t>& destination, std::string_view bytes)
    { destination.insert(destination.end(), reinterpret_cast<const uint8_t*>(bytes.data()), reinterpret_cast<const uint8_t*>(bytes.data() + bytes.size())); };
    appendBytes(snapshotPayload, "\x1b[2J");
    for (uint16_t row = 0; row < kRows; ++row)
    {
        std::array<char, 16> cursorPosition{};
        char* cursor         = cursorPosition.data();
        *cursor++            = '\x1b';
        *cursor++            = '[';
        const auto converted = std::to_chars(cursor, cursorPosition.data() + cursorPosition.size() - 3, static_cast<uint32_t>(row) + 1U);
        if (converted.ec != std::errc{})
        {
            return false;
        }
        cursor    = converted.ptr;
        *cursor++ = ';';
        *cursor++ = '1';
        *cursor++ = 'H';
        appendBytes(snapshotPayload, std::string_view(cursorPosition.data(), static_cast<size_t>(cursor - cursorPosition.data())));
        for (uint16_t column = 0; column < kColumns; ++column)
        {
            snapshotPayload.push_back(static_cast<uint8_t>('a' + ((row + column) % 26)));
        }
    }
    appendBytes(snapshotPayload,
                "\x1b_Ga=T,t=d,f=32,i=1,p=1,s=1,v=1,q=2;"
                "/wAA/w==\x1b\\");
    constexpr std::string_view wrappedLine = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
                                             "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
                                             "wrapped\r\n";
    resizePayload.reserve(wrappedLine.size() * 10'000U);
    for (size_t index = 0; index < 10'000U; ++index)
    {
        appendBytes(resizePayload, wrappedLine);
    }

    size_t contextIndex = 0;
    auto nextContext    = [&]() -> ResourceCallbackContext*
    {
        if (contextIndex >= contexts.size())
        {
            return nullptr;
        }
        return &contexts[contextIndex++];
    };
    auto initializeMeasurement = [&measurements](size_t index, const char* id, const char* unit) -> Measurement&
    {
        measurements[index]          = {};
        measurements[index].metricId = id;
        measurements[index].unit     = unit;
        return measurements[index];
    };

    PerformanceSession active;
    ResourceCallbackContext* selected = nullptr;
    Measurement& create               = initializeMeasurement(0, "create-destroy-us", "microseconds");
    if (! CollectMeasurement(create,
                             10,
                             50,
                             qpcFrequency,
                             []() { return true; },
                             [&]()
    {
        selected = nextContext();
        if (selected == nullptr)
        {
            return false;
        }
        const bool created     = CreatePerformanceSession(adapter, limits, ledger, *selected, 0, "performance-create-destroy", "performance", active);
        const bool snapshotted = created && RunBenchmarkAction(adapter, active.session, "benchmark-full-snapshot", "{\"copy\":\"full\"}");
        const bool destroyed   = created && DestroyPerformanceSession(adapter, ledger, active);
        return created && snapshotted && destroyed;
    },
                             []() { return true; }))
    {
        return false;
    }

    Measurement& parse = initializeMeasurement(1, "parse-8mib-us", "microseconds");
    if (! CollectMeasurement(parse,
                             5,
                             30,
                             qpcFrequency,
                             [&]()
    {
        selected = nextContext();
        return selected != nullptr && CreatePerformanceSession(adapter, limits, ledger, *selected, 0, "performance-parse", "fibonacci", active);
    },
                             [&]()
    {
        constexpr std::array<size_t, 10> chunks = {
            1,
            2,
            3,
            5,
            8,
            13,
            21,
            34,
            55,
            89,
        };
        size_t offset        = 0;
        size_t scheduleIndex = 0;
        while (offset < parsePayload.size())
        {
            const size_t chunk = std::min(chunks[scheduleIndex++ % chunks.size()], parsePayload.size() - offset);
            if (adapter.session_write(active.session, parsePayload.data() + offset, chunk) != RS_TERMINAL_GATE0_SUCCESS)
            {
                return false;
            }
            offset += chunk;
        }
        return true;
    },
                             [&]() { return DestroyPerformanceSession(adapter, ledger, active); }))
    {
        return false;
    }

    auto prepareSnapshot = [&]()
    {
        selected = nextContext();
        return selected != nullptr && CreatePerformanceSession(adapter, limits, ledger, *selected, 0, "performance-snapshot", "performance", active) &&
               adapter.session_write(active.session, snapshotPayload.data(), snapshotPayload.size()) == RS_TERMINAL_GATE0_SUCCESS;
    };
    auto cleanupActive = [&]() { return DestroyPerformanceSession(adapter, ledger, active); };
    Measurement& full  = initializeMeasurement(2, "full-snapshot-us", "microseconds");
    if (! CollectMeasurement(full, 10, 50, qpcFrequency, prepareSnapshot, [&]() {
        return RunBenchmarkAction(adapter, active.session, "benchmark-full-snapshot", "{\"copy\":\"full\"}");
    }, cleanupActive))
    {
        return false;
    }

    Measurement& dirty = initializeMeasurement(3, "dirty-snapshot-us", "microseconds");
    if (! CollectMeasurement(dirty,
                             10,
                             50,
                             qpcFrequency,
                             [&]()
    {
        if (! prepareSnapshot() || ! RunBenchmarkAction(adapter, active.session, "benchmark-full-snapshot", "{\"copy\":\"full\"}"))
        {
            return false;
        }
        constexpr std::string_view mutation = "\x1b[1;1HA\x1b[2;1HB\x1b[3;3H";
        return adapter.session_write(active.session, reinterpret_cast<const uint8_t*>(mutation.data()), mutation.size()) == RS_TERMINAL_GATE0_SUCCESS;
    },
                             [&]() { return RunBenchmarkAction(adapter, active.session, "benchmark-dirty-snapshot", "{\"copy\":\"dirty\"}"); },
                             cleanupActive))
    {
        return false;
    }

    Measurement& resize = initializeMeasurement(4, "resize-reflow-us", "microseconds");
    if (! CollectMeasurement(resize,
                             5,
                             30,
                             qpcFrequency,
                             [&]()
    {
        selected = nextContext();
        return selected != nullptr && CreatePerformanceSession(adapter, limits, ledger, *selected, 0, "performance-resize-reflow", "performance", active) &&
               adapter.session_write(active.session, resizePayload.data(), resizePayload.size()) == RS_TERMINAL_GATE0_SUCCESS;
    },
                             [&]() { return RunBenchmarkAction(adapter, active.session, "benchmark-resize-reflow", "{\"dimensions\":\"80x50,120x40\"}"); },
                             cleanupActive))
    {
        return false;
    }

    Measurement& kitty = initializeMeasurement(5, "kitty-1024-rgba-us", "microseconds");
    if (! CollectMeasurement(kitty,
                             3,
                             20,
                             qpcFrequency,
                             [&]()
    {
        selected = nextContext();
        return selected != nullptr && CreatePerformanceSession(adapter, limits, ledger, *selected, 0, "performance-kitty", "performance", active);
    },
                             [&]()
    { return RunBenchmarkAction(adapter, active.session, "benchmark-kitty-1024-rgba", "{\"format\":\"rgba\",\"height\":\"1024\",\"width\":\"1024\"}"); },
                             cleanupActive))
    {
        return false;
    }

    Measurement& storm = initializeMeasurement(6, "eight-session-storm-us", "microseconds");
    std::array<PerformanceSession, 8> stormSessions{};
    auto stormScenario = [&](uint64_t* retainedCpuBytes)
    {
        stormSessions       = {};
        bool scenarioPassed = true;
        for (size_t index = 0; index < stormSessions.size(); ++index)
        {
            ResourceCallbackContext* context = nextContext();
            if (context == nullptr ||
                ! CreatePerformanceSession(adapter, limits, ledger, *context, index, "performance-storm", "performance", stormSessions[index]))
            {
                scenarioPassed = false;
                break;
            }
        }
        constexpr size_t sessionPayloadBytes = 1024U * 1024U;
        constexpr size_t chunkBytes          = 4096U;
        for (size_t offset = 0; scenarioPassed && offset < sessionPayloadBytes; offset += chunkBytes)
        {
            const size_t length = std::min(chunkBytes, sessionPayloadBytes - offset);
            for (const auto& session : stormSessions)
            {
                if (session.session == nullptr || adapter.session_write(session.session, parsePayload.data() + offset, length) != RS_TERMINAL_GATE0_SUCCESS)
                {
                    scenarioPassed = false;
                    break;
                }
            }
        }
        for (const auto& session : stormSessions)
        {
            if (scenarioPassed && ! RunBenchmarkAction(adapter, session.session, "benchmark-full-snapshot", "{\"copy\":\"full\"}"))
            {
                scenarioPassed = false;
            }
        }
        if (scenarioPassed && retainedCpuBytes != nullptr)
        {
            *retainedCpuBytes = ledger.CurrentCpuBytes();
            if (*retainedCpuBytes == std::numeric_limits<uint64_t>::max())
            {
                scenarioPassed = false;
            }
        }
        for (size_t reverse = stormSessions.size(); reverse != 0; --reverse)
        {
            if (stormSessions[reverse - 1].session != nullptr && ! DestroyPerformanceSession(adapter, ledger, stormSessions[reverse - 1]))
            {
                scenarioPassed = false;
            }
        }
        return scenarioPassed;
    };
    if (! CollectMeasurement(storm, 3, 20, qpcFrequency, []() { return true; }, [&]() { return stormScenario(nullptr); }, []() { return true; }))
    {
        return false;
    }

    Measurement& retained = initializeMeasurement(7, "eight-session-retained-bytes", "bytes");
    retained.warmupCount  = 1;
    retained.sampleCount  = 5;
    for (size_t iteration = 0; iteration < 6; ++iteration)
    {
        uint64_t retainedCpuBytes = 0;
        if (! stormScenario(&retainedCpuBytes))
        {
            return false;
        }
        if (iteration != 0)
        {
            retained.samples[iteration - 1] = retainedCpuBytes;
        }
    }
    Measurement& package      = initializeMeasurement(8, "added-release-package-bytes", "bytes");
    package.sampleCount       = 1;
    package.samples[0]        = arguments.addedReleasePackageBytes;
    Measurement& abi          = initializeMeasurement(9, "crossed-abi-elements", "count");
    abi.sampleCount           = 1;
    abi.samples[0]            = arguments.crossedAbiElements;
    Measurement& dependencies = initializeMeasurement(10, "non-system-runtime-dependencies", "count");
    dependencies.sampleCount  = 1;
    dependencies.samples[0]   = arguments.nonSystemRuntimeDependencies;
    return contextIndex <= contexts.size() && ledger.CheckedArithmetic() && ledger.PeaksWithinLimits() && ledger.AllReleased();
}

struct HarnessCheck
{
    const char* id             = "";
    bool passed                = false;
    const char* resultCategory = "";
};

struct HarnessResult
{
    std::array<HarnessCheck, 8> checks{};
    size_t checkCount = 0;
    std::vector<FixtureRow> rows;
    std::array<Measurement, 11> measurements{};
    size_t measurementCount = 0;
    uint64_t qpcFrequency   = 0;
    bool passed             = true;

    void AddCheck(const char* id, bool value, const char* passCategory, const char* failCategory) noexcept
    {
        if (checkCount < checks.size())
        {
            checks[checkCount++] = {
                .id             = id,
                .passed         = value,
                .resultCategory = value ? passCategory : failCategory,
            };
        }
        passed = passed && value;
    }
};

void PrintJsonString(std::string_view value) noexcept
{
    std::fputc('"', stdout);
    for (const unsigned char character : value)
    {
        switch (character)
        {
            case '"': std::fputs("\\\"", stdout); break;
            case '\\': std::fputs("\\\\", stdout); break;
            case '\b': std::fputs("\\b", stdout); break;
            case '\f': std::fputs("\\f", stdout); break;
            case '\n': std::fputs("\\n", stdout); break;
            case '\r': std::fputs("\\r", stdout); break;
            case '\t': std::fputs("\\t", stdout); break;
            default:
                if (character < 0x20)
                {
                    std::fprintf(stdout, "\\u%04x", character);
                }
                else
                {
                    std::fputc(character, stdout);
                }
                break;
        }
    }
    std::fputc('"', stdout);
}

void EmitResult(const Arguments& arguments, const HarnessResult& result, bool adapterIdentityAvailable) noexcept
{
    std::fputs("{\"formatId\":\"red-salamander-terminal-engine-gate0-harness\","
               "\"version\":1,\"status\":\"",
               stdout);
    std::fputs(result.passed ? "pass" : "fail", stdout);
    std::fputs("\",\"architecture\":", stdout);
    PrintJsonString(kArchitectureName);
    std::fputs(",\"measurementMode\":", stdout);
    PrintJsonString(arguments.fullMeasurements ? "full" : "functional");
    std::fputs(",\"fixtureManifestSha256\":", stdout);
    PrintJsonString(arguments.expectedFixtureManifestSha256);
    std::fputs(",\"adapterIdentity\":", stdout);
    if (! adapterIdentityAvailable)
    {
        std::fputs("null", stdout);
    }
    else
    {
        std::fputs("{\"candidateId\":", stdout);
        PrintJsonString(arguments.expectedCandidateId);
        std::fputs(",\"upstreamPin\":", stdout);
        PrintJsonString(arguments.expectedUpstreamPin);
        std::fputs(",\"sourceSha256\":", stdout);
        PrintJsonString(arguments.expectedSourceSha256);
        std::fputs(",\"adapterHeaderSha256\":", stdout);
        PrintJsonString(arguments.expectedAdapterHeaderSha256);
        std::fputs(",\"buildIdentity\":", stdout);
        PrintJsonString(arguments.expectedBuildIdentity);
        std::fprintf(stdout,
                     ",\"adapterAbi\":%u,\"capabilities\":\"%016llx\"}",
                     arguments.expectedAdapterAbi,
                     static_cast<unsigned long long>(arguments.expectedCapabilities));
    }
    std::fputs(",\"checks\":[", stdout);
    for (size_t index = 0; index < result.checkCount; ++index)
    {
        if (index != 0)
        {
            std::fputc(',', stdout);
        }
        const auto& check = result.checks[index];
        std::fputs("{\"id\":", stdout);
        PrintJsonString(check.id);
        std::fputs(",\"status\":\"", stdout);
        std::fputs(check.passed ? "pass" : "fail", stdout);
        std::fputs("\",\"resultCategory\":", stdout);
        PrintJsonString(check.resultCategory);
        std::fputc('}', stdout);
    }
    std::fputs("],\"fixtureRows\":[", stdout);
    for (size_t index = 0; index < result.rows.size(); ++index)
    {
        if (index != 0)
        {
            std::fputc(',', stdout);
        }
        const auto& row = result.rows[index];
        std::fputs("{\"fixtureId\":", stdout);
        PrintJsonString(row.fixtureId);
        std::fputs(",\"scheduleId\":", stdout);
        PrintJsonString(row.scheduleId);
        std::fputs(",\"inputSha256\":", stdout);
        PrintJsonString(row.inputSha256);
        std::fprintf(stdout, ",\"inputSizeBytes\":\"%llu\",\"expectedStateSha256\":", static_cast<unsigned long long>(row.inputSizeBytes));
        PrintJsonString(row.expectedStateSha256);
        std::fputs(",\"actualStateSha256\":", stdout);
        PrintJsonString(row.actualStateSha256);
        std::fputs(",\"expectedResourceTraceSha256\":", stdout);
        PrintJsonString(row.expectedResourceTraceSha256);
        std::fputs(",\"actualResourceTraceSha256\":", stdout);
        PrintJsonString(row.actualResourceTraceSha256);
        std::fputs(",\"status\":\"", stdout);
        std::fputs(row.passed ? "pass" : "fail", stdout);
        std::fputs("\",\"resultCategory\":", stdout);
        PrintJsonString(row.resultCategory);
        std::fputc('}', stdout);
    }
    std::fprintf(stdout, "],\"qpcFrequency\":\"%llu\",\"measurements\":[", static_cast<unsigned long long>(result.qpcFrequency));
    for (size_t index = 0; index < result.measurementCount; ++index)
    {
        if (index != 0)
        {
            std::fputc(',', stdout);
        }
        const auto& measurement = result.measurements[index];
        std::fputs("{\"metricId\":", stdout);
        PrintJsonString(measurement.metricId);
        std::fputs(",\"unit\":", stdout);
        PrintJsonString(measurement.unit);
        std::fprintf(stdout, ",\"warmupCount\":%u,\"samples\":[", measurement.warmupCount);
        for (size_t sample = 0; sample < measurement.sampleCount; ++sample)
        {
            if (sample != 0)
            {
                std::fputc(',', stdout);
            }
            std::fprintf(stdout, "\"%llu\"", static_cast<unsigned long long>(measurement.samples[sample]));
        }
        std::fputs("]}", stdout);
    }
    std::fputs("],\"limitations\":[]}\n", stdout);
}

} // namespace

int main(int argc, char** argv)
{
    Arguments arguments;
    if (! ParseArguments(argc, argv, arguments))
    {
        std::fputs("Invalid arguments for TerminalEngineGate0Harness.\n", stderr);
        return 2;
    }

    HarnessResult result;
    const bool fixtureContract = VerifyFrozenFixtureShape(arguments);
    result.AddCheck("frozen-fixture-contract", fixtureContract, "fixture-contract-matched", "fixture-contract-mismatch");
    if (! fixtureContract)
    {
        EmitResult(arguments, result, false);
        return 12;
    }

    LoadedAdapter loaded;
    const bool adapterIdentity = LoadAndVerifyAdapter(arguments, loaded);
    result.AddCheck("adapter-runtime-identity", adapterIdentity, "adapter-identity-matched", "adapter-identity-mismatch");
    if (! adapterIdentity)
    {
        EmitResult(arguments, result, false);
        return 10;
    }

    constexpr auto limits = MakeFrozenResourceLimits();
    ProcessResourceLedger ledger(limits);
    std::array<ResourceCallbackContext, 768> callbackContexts{};
    const bool fixtureExecution = RunFixtureMatrix(loaded.api, limits, ledger, std::span(callbackContexts.data(), kMaximumFixtureRows), result.rows);
    result.AddCheck("fixture-matrix", fixtureExecution, "every-fixture-schedule-action-matched", "fixture-schedule-action-mismatch");

    const bool sessionStorm = RunEightSessionStorm(loaded.api, limits, ledger, std::span(callbackContexts.data() + kMaximumFixtureRows, 8));
    result.AddCheck("shared-process-ledger", sessionStorm, "eight-session-ledger-balanced", "eight-session-ledger-failed");
    if (arguments.fullMeasurements)
    {
        const bool measurements =
            RunFullMeasurements(arguments,
                                loaded.api,
                                limits,
                                ledger,
                                std::span(callbackContexts.data() + kMaximumFixtureRows + 8, callbackContexts.size() - kMaximumFixtureRows - 8),
                                result.measurements,
                                result.qpcFrequency);
        result.measurementCount = measurements ? result.measurements.size() : 0;
        result.AddCheck("full-measurements", measurements, "all-eleven-measurements-collected", "neutral-performance-runner-failed");
    }

    bool callbacksQuiet = true;
    for (const auto& context : callbackContexts)
    {
        callbacksQuiet = callbacksQuiet && ! context.callbackError.load(std::memory_order_acquire);
    }
    result.AddCheck("callback-quiet-point", callbacksQuiet, "callbacks-quiet-before-unload", "callback-after-session-destroy");

    HMODULE module      = loaded.module.release();
    const bool unloaded = module != nullptr && FreeLibrary(module) != FALSE;
    result.AddCheck("adapter-unload", unloaded, "adapter-unloaded-after-quiet-point", "adapter-unload-failed");

    EmitResult(arguments, result, true);
    return result.passed ? 0 : 11;
}

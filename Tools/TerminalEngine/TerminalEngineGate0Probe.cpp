#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <wincrypt.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#include "TerminalEngineGate0Layout.h"
#include "Generated/TerminalEngineFixtures.v1.h"

#include <algorithm>
#include <array>
#include <bit>
#include <charconv>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <limits>
#include <span>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

namespace
{
namespace FixtureContract =
    RedSalamander::TerminalEngine::Fixtures::V1;

constexpr size_t kMaximumStateBytes = 1024U * 1024U;
constexpr size_t kMaximumCallbackEvents = 4096;
constexpr uint64_t kMaximumDenialOrdinals = 128;
constexpr size_t kDefaultRandomCases = 24;
constexpr size_t kMaximumRandomCases = 64;
constexpr size_t kDefaultMaximumCaseBytes = 256;
constexpr size_t kMaximumCaseBytes = 1024;
constexpr uint64_t kMaximumLayoutJcsBytes = 64U * 1024U;
constexpr size_t kMaximumC7StateBytes = 8192;
constexpr size_t kC7FixedStreamBytes = 192;
constexpr uint64_t kDefaultSeed = UINT64_C(0x5253474154453031);
constexpr std::string_view kFuzzFixtureId = "gate0-public-fuzz-v1";
constexpr std::string_view kFuzzScheduleId = "deterministic-seeded-v1";
constexpr std::string_view kC7FixtureId = "c7-engine-effects-v1";
constexpr std::string_view kC7ActionId = "probe-c7-engine-effects";
constexpr std::string_view kC7ActionArguments =
    "{\"contractId\":\"c7-engine-effects-v1\"}";
constexpr uint64_t kRequiredCapabilities =
    RS_TERMINAL_GATE0_CAP_CELLS_AND_STYLES |
    RS_TERMINAL_GATE0_CAP_SNAPSHOT_COPY |
    RS_TERMINAL_GATE0_CAP_INPUT_ENCODING |
    RS_TERMINAL_GATE0_CAP_EFFECT_ORDERING |
    RS_TERMINAL_GATE0_CAP_KITTY_STATIC |
    RS_TERMINAL_GATE0_CAP_RESOURCE_ACCOUNTING |
    RS_TERMINAL_GATE0_CAP_QUIET_DESTROY;

#if defined(_M_X64)
constexpr RsTerminalGate0Architecture kExpectedArchitecture =
    RS_TERMINAL_GATE0_ARCH_X64;
constexpr std::string_view kArchitectureName = "x64";
#elif defined(_M_ARM64)
constexpr RsTerminalGate0Architecture kExpectedArchitecture =
    RS_TERMINAL_GATE0_ARCH_ARM64;
constexpr std::string_view kArchitectureName = "ARM64";
#else
#error The terminal Gate-0 probe supports only x64 and ARM64.
#endif

enum class Mode
{
    Layout,
    Fuzz,
    C7,
};

struct Arguments
{
    Mode mode = Mode::Layout;
    std::wstring adapterDll;
    std::string expectedLayoutIdentity;
    std::string expectedCandidateId;
    std::string expectedUpstreamPin;
    std::string expectedSourceSha256;
    std::string expectedAdapterHeaderSha256;
    std::string expectedBuildIdentity;
    std::string expectedArchitecture;
    std::string probeLayoutJcs;
    std::string probeLayoutSha256;
    std::vector<uint8_t> c7Stream;
    uint64_t seed = kDefaultSeed;
    uint64_t expectedCapabilities = 0;
    size_t randomCases = kDefaultRandomCases;
    size_t maximumCaseBytes = kDefaultMaximumCaseBytes;
    bool hasMode = false;
    bool hasExpectedCapabilities = false;
};

void AppendUnsigned(std::string& output, uint64_t value)
{
    std::array<char, 32> buffer{};
    const auto result = std::to_chars(
        buffer.data(),
        buffer.data() + buffer.size(),
        value);
    if (result.ec == std::errc{})
    {
        output.append(buffer.data(), result.ptr);
    }
}

void AppendJcsString(std::string& output, std::string_view value)
{
    constexpr char hex[] = "0123456789abcdef";
    output.push_back('"');
    for (const unsigned char character : value)
    {
        switch (character)
        {
        case '"':
            output.append("\\\"");
            break;
        case '\\':
            output.append("\\\\");
            break;
        case '\b':
            output.append("\\b");
            break;
        case '\f':
            output.append("\\f");
            break;
        case '\n':
            output.append("\\n");
            break;
        case '\r':
            output.append("\\r");
            break;
        case '\t':
            output.append("\\t");
            break;
        default:
            if (character < 0x20)
            {
                output.append("\\u00");
                output.push_back(hex[character >> 4]);
                output.push_back(hex[character & 0x0F]);
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

[[nodiscard]] bool Sha256Hex(
    std::span<const uint8_t> bytes,
    std::string& output) noexcept
{
    if (bytes.size() >
        static_cast<size_t>(std::numeric_limits<DWORD>::max()))
    {
        return false;
    }
    std::array<uint8_t, 32> digest{};
    DWORD digestSize = static_cast<DWORD>(digest.size());
    if (!CryptHashCertificate2(
            L"SHA256",
            0,
            nullptr,
            bytes.data(),
            static_cast<DWORD>(bytes.size()),
            digest.data(),
            &digestSize) ||
        digestSize != digest.size())
    {
        return false;
    }
    constexpr char hex[] = "0123456789abcdef";
    output.clear();
    output.reserve(digest.size() * 2);
    for (const uint8_t byte : digest)
    {
        output.push_back(hex[byte >> 4]);
        output.push_back(hex[byte & 0x0F]);
    }
    return true;
}

[[nodiscard]] bool Sha256Hex(
    std::string_view value,
    std::string& output) noexcept
{
    return Sha256Hex(
        std::span(
            reinterpret_cast<const uint8_t*>(value.data()),
            value.size()),
        output);
}

[[nodiscard]] std::string BuildLayoutJcs()
{
    return RedSalamander::TerminalEngine::Gate0::Layout::
        BuildLayoutJcs();
}

[[nodiscard]] bool IsLowerHex(
    std::string_view value,
    size_t count) noexcept
{
    if (value.size() != count)
    {
        return false;
    }
    return std::ranges::all_of(value, [](const char character) {
        return (character >= '0' && character <= '9') ||
            (character >= 'a' && character <= 'f');
    });
}

[[nodiscard]] bool IsLayoutIdentity(std::string_view value) noexcept
{
    constexpr std::string_view prefix = "layout-v1-";
    return value.starts_with(prefix) &&
        IsLowerHex(value.substr(prefix.size()), 64);
}

[[nodiscard]] bool IsBuildIdentity(std::string_view value) noexcept
{
    constexpr std::string_view prefix = "build-v1-";
    return value.starts_with(prefix) &&
        IsLowerHex(value.substr(prefix.size()), 64);
}

[[nodiscard]] bool IsBoundedPrintableAscii(
    std::string_view value,
    size_t maximumLength) noexcept
{
    return !value.empty() &&
        value.size() <= maximumLength &&
        std::ranges::all_of(value, [](const unsigned char character) {
            return character >= 0x21 && character <= 0x7E;
        });
}

[[nodiscard]] bool ParseUnsigned(
    std::wstring_view text,
    uint64_t& value) noexcept
{
    if (text.empty())
    {
        return false;
    }
    std::string narrow;
    narrow.reserve(text.size());
    for (const wchar_t character : text)
    {
        if (character > 0x7F)
        {
            return false;
        }
        narrow.push_back(static_cast<char>(character));
    }
    int base = 10;
    std::string_view digits = narrow;
    if (digits.starts_with("0x") || digits.starts_with("0X"))
    {
        base = 16;
        digits.remove_prefix(2);
    }
    if (digits.empty())
    {
        return false;
    }
    const auto result = std::from_chars(
        digits.data(),
        digits.data() + digits.size(),
        value,
        base);
    return result.ec == std::errc{} &&
        result.ptr == digits.data() + digits.size();
}

[[nodiscard]] bool NarrowAscii(
    std::wstring_view value,
    std::string& output)
{
    output.clear();
    output.reserve(value.size());
    for (const wchar_t character : value)
    {
        if (character > 0x7F)
        {
            return false;
        }
        output.push_back(static_cast<char>(character));
    }
    return true;
}

[[nodiscard]] bool DecodeLowerHex(
    std::wstring_view value,
    std::vector<uint8_t>& output)
{
    if ((value.size() % 2) != 0)
    {
        return false;
    }
    output.clear();
    output.reserve(value.size() / 2);
    const auto nibble = [](wchar_t character) -> int {
        if (character >= L'0' && character <= L'9')
        {
            return character - L'0';
        }
        if (character >= L'a' && character <= L'f')
        {
            return character - L'a' + 10;
        }
        return -1;
    };
    for (size_t index = 0; index < value.size(); index += 2)
    {
        const int high = nibble(value[index]);
        const int low = nibble(value[index + 1]);
        if (high < 0 || low < 0)
        {
            output.clear();
            return false;
        }
        output.push_back(
            static_cast<uint8_t>((high << 4) | low));
    }
    return true;
}

[[nodiscard]] bool ParseArguments(
    int argc,
    wchar_t** argv,
    Arguments& arguments)
{
    if (argc < 3 || (argc % 2) == 0)
    {
        return false;
    }
    for (int index = 1; index < argc; index += 2)
    {
        const std::wstring_view name(argv[index]);
        const std::wstring_view value(argv[index + 1]);
        if (name == L"--mode")
        {
            if (value == L"layout")
            {
                arguments.mode = Mode::Layout;
            }
            else if (value == L"fuzz")
            {
                arguments.mode = Mode::Fuzz;
            }
            else if (value == L"c7")
            {
                arguments.mode = Mode::C7;
            }
            else
            {
                return false;
            }
            arguments.hasMode = true;
        }
        else if (name == L"--adapter-dll")
        {
            arguments.adapterDll.assign(value);
        }
        else if (name == L"--expected-layout-identity")
        {
            if (!NarrowAscii(
                    value,
                    arguments.expectedLayoutIdentity))
            {
                return false;
            }
        }
        else if (name == L"--expected-candidate-id")
        {
            if (!NarrowAscii(value, arguments.expectedCandidateId))
            {
                return false;
            }
        }
        else if (name == L"--expected-upstream-pin")
        {
            if (!NarrowAscii(value, arguments.expectedUpstreamPin))
            {
                return false;
            }
        }
        else if (name == L"--expected-source-sha256")
        {
            if (!NarrowAscii(value, arguments.expectedSourceSha256))
            {
                return false;
            }
        }
        else if (name == L"--expected-adapter-header-sha256")
        {
            if (!NarrowAscii(
                    value,
                    arguments.expectedAdapterHeaderSha256))
            {
                return false;
            }
        }
        else if (name == L"--expected-build-identity")
        {
            if (!NarrowAscii(value, arguments.expectedBuildIdentity))
            {
                return false;
            }
        }
        else if (name == L"--expected-architecture")
        {
            if (!NarrowAscii(value, arguments.expectedArchitecture))
            {
                return false;
            }
        }
        else if (name == L"--expected-capabilities")
        {
            if (!ParseUnsigned(
                    value,
                    arguments.expectedCapabilities))
            {
                return false;
            }
            arguments.hasExpectedCapabilities = true;
        }
        else if (name == L"--c7-stream-hex")
        {
            if (!DecodeLowerHex(value, arguments.c7Stream))
            {
                return false;
            }
        }
        else if (name == L"--seed")
        {
            if (!ParseUnsigned(value, arguments.seed))
            {
                return false;
            }
        }
        else if (name == L"--random-cases")
        {
            uint64_t parsed = 0;
            if (!ParseUnsigned(value, parsed) ||
                parsed == 0 ||
                parsed > kMaximumRandomCases)
            {
                return false;
            }
            arguments.randomCases = static_cast<size_t>(parsed);
        }
        else if (name == L"--maximum-case-bytes")
        {
            uint64_t parsed = 0;
            if (!ParseUnsigned(value, parsed) ||
                parsed == 0 ||
                parsed > kMaximumCaseBytes)
            {
                return false;
            }
            arguments.maximumCaseBytes =
                static_cast<size_t>(parsed);
        }
        else
        {
            return false;
        }
    }
    if (!arguments.hasMode)
    {
        return false;
    }
    if (arguments.mode == Mode::Layout)
    {
        return arguments.adapterDll.empty() &&
            arguments.expectedLayoutIdentity.empty() &&
            arguments.expectedCandidateId.empty() &&
            arguments.expectedUpstreamPin.empty() &&
            arguments.expectedSourceSha256.empty() &&
            arguments.expectedAdapterHeaderSha256.empty() &&
            arguments.expectedBuildIdentity.empty() &&
            arguments.expectedArchitecture.empty() &&
            !arguments.hasExpectedCapabilities &&
            arguments.c7Stream.empty();
    }
    const bool identityValid =
        !arguments.adapterDll.empty() &&
        IsLayoutIdentity(arguments.expectedLayoutIdentity) &&
        IsBoundedPrintableAscii(
            arguments.expectedCandidateId,
            256) &&
        IsBoundedPrintableAscii(
            arguments.expectedUpstreamPin,
            256) &&
        IsLowerHex(arguments.expectedSourceSha256, 64) &&
        IsLowerHex(
            arguments.expectedAdapterHeaderSha256,
            64) &&
        IsBuildIdentity(arguments.expectedBuildIdentity) &&
        arguments.expectedArchitecture == kArchitectureName &&
        arguments.hasExpectedCapabilities &&
        arguments.expectedCapabilities == kRequiredCapabilities;
    if (!identityValid)
    {
        return false;
    }
    if (arguments.mode == Mode::C7)
    {
        return arguments.c7Stream.size() ==
            kC7FixedStreamBytes;
    }
    return arguments.c7Stream.empty();
}

[[nodiscard]] RsTerminalGate0String AdapterString(
    std::string_view value) noexcept
{
    return {
        .ptr = reinterpret_cast<const uint8_t*>(value.data()),
        .len = value.size(),
    };
}

[[nodiscard]] std::string_view AdapterStringView(
    RsTerminalGate0String value) noexcept
{
    if (value.ptr == nullptr)
    {
        return {};
    }
    return {
        reinterpret_cast<const char*>(value.ptr),
        value.len,
    };
}

[[nodiscard]] bool ValidAdapterString(
    RsTerminalGate0String value,
    size_t maximumLength) noexcept
{
    return value.len <= maximumLength &&
        (value.len == 0 || value.ptr != nullptr);
}

[[nodiscard]] constexpr RsTerminalGate0ResourceLimitsV1
MakeFuzzResourceLimits() noexcept
{
    return {
        .size = sizeof(RsTerminalGate0ResourceLimitsV1),
        .terminal_text_session_bytes = 64U * 1024U * 1024U,
        .terminal_text_process_bytes = 512U * 1024U * 1024U,
        .image_cpu_session_bytes = 64U * 1024U * 1024U,
        .image_cpu_process_bytes = 256U * 1024U * 1024U,
        .image_gpu_session_bytes = 64U * 1024U * 1024U,
        .image_gpu_process_bytes = 256U * 1024U * 1024U,
        .pending_effect_session_bytes = 2U * 1024U * 1024U,
        .pending_effect_process_bytes = 32U * 2U * 1024U * 1024U,
        .pending_effect_session_descriptors = 256,
        .pending_effect_process_descriptors = 32U * 256U,
        .maximum_sessions = 32,
        .kitty_max_pixels = 16'000'000,
        .kitty_max_dimension = 10'000,
        .reserved = 0,
    };
}

struct CallbackEvent
{
    uint64_t ordinal = 0;
    uint64_t bytes = 0;
    uint32_t descriptors = 0;
    uint8_t kind = 0;
    bool reserve = false;
    bool accepted = false;
};

class ResourceLedger final
{
public:
    ResourceLedger(
        const RsTerminalGate0ResourceLimitsV1& limits,
        uint64_t denialOrdinal) noexcept
        : _limits(limits),
          _denialOrdinal(denialOrdinal)
    {
    }

    [[nodiscard]] bool Reserve(
        RsTerminalGate0ResourceKind kind,
        uint64_t bytes,
        uint32_t descriptors) noexcept
    {
        AcquireSRWLockExclusive(&_lock);
        const auto unlock = wil::scope_exit(
            [&] { ReleaseSRWLockExclusive(&_lock); });
        ++_reserveCalls;
        bool accepted =
            !_closed &&
            static_cast<uint32_t>(kind) <
                RS_TERMINAL_GATE0_RESOURCE_COUNT &&
            _eventCount < _events.size() &&
            (_denialOrdinal == 0 ||
             _reserveCalls != _denialOrdinal);
        if (accepted)
        {
            const size_t index = static_cast<size_t>(kind);
            const uint64_t byteLimit = ByteLimit(kind);
            const uint32_t descriptorLimit =
                DescriptorLimit(kind);
            accepted =
                bytes <= byteLimit &&
                _bytes[index] <= byteLimit - bytes &&
                descriptors <= descriptorLimit &&
                _descriptors[index] <=
                    descriptorLimit - descriptors;
            if (accepted)
            {
                _bytes[index] += bytes;
                _descriptors[index] += descriptors;
            }
        }
        AddEvent(kind, bytes, descriptors, true, accepted);
        if (!accepted &&
            _denialOrdinal != 0 &&
            _reserveCalls == _denialOrdinal)
        {
            _denialObserved = true;
        }
        return accepted;
    }

    void Release(
        RsTerminalGate0ResourceKind kind,
        uint64_t bytes,
        uint32_t descriptors) noexcept
    {
        AcquireSRWLockExclusive(&_lock);
        const auto unlock = wil::scope_exit(
            [&] { ReleaseSRWLockExclusive(&_lock); });
        bool accepted =
            !_closed &&
            static_cast<uint32_t>(kind) <
                RS_TERMINAL_GATE0_RESOURCE_COUNT &&
            _eventCount < _events.size();
        if (accepted)
        {
            const size_t index = static_cast<size_t>(kind);
            accepted =
                bytes <= _bytes[index] &&
                descriptors <= _descriptors[index];
            if (accepted)
            {
                _bytes[index] -= bytes;
                _descriptors[index] -= descriptors;
            }
        }
        if (!accepted)
        {
            _error = true;
        }
        AddEvent(kind, bytes, descriptors, false, accepted);
    }

    void Close() noexcept
    {
        AcquireSRWLockExclusive(&_lock);
        const auto unlock = wil::scope_exit(
            [&] { ReleaseSRWLockExclusive(&_lock); });
        _closed = true;
    }

    [[nodiscard]] bool IsQuiet() noexcept
    {
        AcquireSRWLockShared(&_lock);
        const auto unlock = wil::scope_exit(
            [&] { ReleaseSRWLockShared(&_lock); });
        return !_error &&
            std::ranges::all_of(
                _bytes,
                [](uint64_t value) { return value == 0; }) &&
            std::ranges::all_of(
                _descriptors,
                [](uint32_t value) { return value == 0; });
    }

    [[nodiscard]] uint64_t ReserveCalls() noexcept
    {
        AcquireSRWLockShared(&_lock);
        const auto unlock = wil::scope_exit(
            [&] { ReleaseSRWLockShared(&_lock); });
        return _reserveCalls;
    }

    [[nodiscard]] bool DenialObserved() noexcept
    {
        AcquireSRWLockShared(&_lock);
        const auto unlock = wil::scope_exit(
            [&] { ReleaseSRWLockShared(&_lock); });
        return _denialObserved;
    }

    [[nodiscard]] std::string CanonicalTrace()
    {
        AcquireSRWLockShared(&_lock);
        const auto unlock = wil::scope_exit(
            [&] { ReleaseSRWLockShared(&_lock); });
        std::string output("[");
        output.reserve(_eventCount * 112);
        for (size_t index = 0; index < _eventCount; ++index)
        {
            if (index != 0)
            {
                output.push_back(',');
            }
            const CallbackEvent& event = _events[index];
            output.append("{\"accepted\":");
            output.append(event.accepted ? "true" : "false");
            output.append(",\"bytes\":");
            AppendUnsigned(output, event.bytes);
            output.append(",\"descriptors\":");
            AppendUnsigned(output, event.descriptors);
            output.append(",\"kind\":");
            AppendUnsigned(output, event.kind);
            output.append(",\"operation\":");
            AppendJcsString(
                output,
                event.reserve ? "reserve" : "release");
            output.append(",\"ordinal\":");
            AppendUnsigned(output, event.ordinal);
            output.push_back('}');
        }
        output.push_back(']');
        return output;
    }

private:
    [[nodiscard]] uint64_t ByteLimit(
        RsTerminalGate0ResourceKind kind) const noexcept
    {
        switch (kind)
        {
        case RS_TERMINAL_GATE0_RESOURCE_TERMINAL_TEXT_CPU:
            return _limits.terminal_text_session_bytes;
        case RS_TERMINAL_GATE0_RESOURCE_IMAGE_CPU:
            return _limits.image_cpu_session_bytes;
        case RS_TERMINAL_GATE0_RESOURCE_IMAGE_GPU:
            return _limits.image_gpu_session_bytes;
        case RS_TERMINAL_GATE0_RESOURCE_PENDING_EFFECT:
            return _limits.pending_effect_session_bytes;
        default:
            return 0;
        }
    }

    [[nodiscard]] uint32_t DescriptorLimit(
        RsTerminalGate0ResourceKind kind) const noexcept
    {
        return kind ==
                RS_TERMINAL_GATE0_RESOURCE_PENDING_EFFECT
            ? _limits.pending_effect_session_descriptors
            : 0;
    }

    void AddEvent(
        RsTerminalGate0ResourceKind kind,
        uint64_t bytes,
        uint32_t descriptors,
        bool reserve,
        bool accepted) noexcept
    {
        if (_eventCount >= _events.size())
        {
            _error = true;
            return;
        }
        _events[_eventCount] = {
            .ordinal = _eventCount + 1,
            .bytes = bytes,
            .descriptors = descriptors,
            .kind = static_cast<uint8_t>(kind),
            .reserve = reserve,
            .accepted = accepted,
        };
        ++_eventCount;
    }

    const RsTerminalGate0ResourceLimitsV1& _limits;
    uint64_t _denialOrdinal = 0;
    SRWLOCK _lock = SRWLOCK_INIT;
    std::array<CallbackEvent, kMaximumCallbackEvents> _events{};
    std::array<uint64_t, RS_TERMINAL_GATE0_RESOURCE_COUNT> _bytes{};
    std::array<uint32_t, RS_TERMINAL_GATE0_RESOURCE_COUNT>
        _descriptors{};
    size_t _eventCount = 0;
    uint64_t _reserveCalls = 0;
    bool _denialObserved = false;
    bool _closed = false;
    bool _error = false;
};

[[nodiscard]] bool RS_TERMINAL_GATE0_CALL ReserveResource(
    void* userdata,
    RsTerminalGate0ResourceKind kind,
    uint64_t bytes,
    uint32_t descriptors) noexcept
{
    if (userdata == nullptr)
    {
        return false;
    }
    return static_cast<ResourceLedger*>(userdata)->Reserve(
        kind,
        bytes,
        descriptors);
}

void RS_TERMINAL_GATE0_CALL ReleaseResource(
    void* userdata,
    RsTerminalGate0ResourceKind kind,
    uint64_t bytes,
    uint32_t descriptors) noexcept
{
    if (userdata != nullptr)
    {
        static_cast<ResourceLedger*>(userdata)->Release(
            kind,
            bytes,
            descriptors);
    }
}

class DeterministicGenerator final
{
public:
    explicit DeterministicGenerator(uint64_t seed) noexcept
        : _state(seed == 0 ? kDefaultSeed : seed)
    {
    }

    [[nodiscard]] uint64_t Next() noexcept
    {
        uint64_t value = _state;
        value ^= value >> 12;
        value ^= value << 25;
        value ^= value >> 27;
        _state = value;
        return value * UINT64_C(2685821657736338717);
    }

private:
    uint64_t _state;
};

struct CorpusCase
{
    std::vector<uint8_t> bytes;
    std::vector<size_t> chunks;
};

[[nodiscard]] std::vector<CorpusCase> BuildCorpus(
    const Arguments& arguments)
{
    std::vector<CorpusCase> corpus;
    corpus.reserve(arguments.randomCases + 5);
    corpus.push_back({.bytes = {}, .chunks = {0}});

    const std::vector<uint8_t> malformed = {
        0xC0, 0xAF, 0xE0, 0x80, 0x80, 0xF5, 0x80, 0x80,
        0x80, 0xFF, 0x1B, 0x5B, 0x3F, 0x1B, 0x5D, 0x38,
        0x3B, 0x3B, 0xC3,
    };
    corpus.push_back(
        {.bytes = malformed, .chunks = {malformed.size()}});
    corpus.push_back(
        {.bytes = malformed,
         .chunks = std::vector<size_t>(malformed.size(), 1)});

    const std::vector<uint8_t> splitControl = {
        0x1B, 0x5F, 0x47, 0x61, 0x3D, 0x74, 0x2C, 0x66,
        0x3D, 0x33, 0x32, 0x3B, 0x41, 0x41, 0x41, 0x1B,
        0x5C, 0x1B, 0x5D, 0x35, 0x32, 0x3B, 0x63, 0x3B,
    };
    corpus.push_back(
        {.bytes = splitControl,
         .chunks = {1, 2, 1, 3, 5, 1, 4, 2, 5}});
    corpus.push_back(
        {.bytes = {0x00, 0x7F, 0x80, 0xBF, 0xFE, 0xFF},
         .chunks = {2, 1, 3}});

    DeterministicGenerator generator(arguments.seed);
    for (size_t caseIndex = 0;
         caseIndex < arguments.randomCases;
         ++caseIndex)
    {
        CorpusCase item;
        const size_t length =
            static_cast<size_t>(
                generator.Next() %
                arguments.maximumCaseBytes) +
            1;
        item.bytes.resize(length);
        for (uint8_t& byte : item.bytes)
        {
            byte = static_cast<uint8_t>(generator.Next());
        }
        size_t remaining = length;
        while (remaining != 0)
        {
            size_t chunk = 0;
            if ((caseIndex % 4) == 0)
            {
                chunk = remaining;
            }
            else if ((caseIndex % 4) == 1)
            {
                chunk = 1;
            }
            else
            {
                chunk = static_cast<size_t>(
                    generator.Next() %
                    std::min<size_t>(remaining, 17)) +
                    1;
            }
            item.chunks.push_back(chunk);
            remaining -= chunk;
        }
        corpus.push_back(std::move(item));
    }
    return corpus;
}

[[nodiscard]] CorpusCase BuildDenialCase(
    const std::vector<CorpusCase>& corpus)
{
    CorpusCase combined;
    size_t total = 0;
    for (const CorpusCase& item : corpus)
    {
        total += item.bytes.size();
    }
    combined.bytes.reserve(total);
    for (const CorpusCase& item : corpus)
    {
        combined.bytes.insert(
            combined.bytes.end(),
            item.bytes.begin(),
            item.bytes.end());
        for (const size_t chunk : item.chunks)
        {
            if (chunk != 0)
            {
                combined.chunks.push_back(chunk);
            }
        }
    }
    if (combined.chunks.empty())
    {
        combined.chunks.push_back(0);
    }
    return combined;
}

[[nodiscard]] bool IsKnownResult(
    RsTerminalGate0Result result) noexcept
{
    return result >= RS_TERMINAL_GATE0_SUCCESS &&
        result <= RS_TERMINAL_GATE0_INTERNAL_ERROR;
}

[[nodiscard]] bool IsSafeWriteResult(
    RsTerminalGate0Result result,
    bool denialEnabled) noexcept
{
    return result == RS_TERMINAL_GATE0_SUCCESS ||
        (denialEnabled &&
         (result == RS_TERMINAL_GATE0_OUT_OF_MEMORY ||
          result == RS_TERMINAL_GATE0_LIMIT_REJECTED));
}

struct TraceResult
{
    std::string adapterIdentitySha256;
    std::string adapterLayoutSha256;
    std::string resourceCanonical;
    std::string resultCanonical;
    std::string stateSha256;
    uint64_t reserveCalls = 0;
    bool denialObserved = false;
    bool unloaded = false;
    bool passed = false;
};

[[nodiscard]] bool AppendAdapterIdentity(
    const RsTerminalGate0IdentityV1& identity,
    std::string& canonical)
{
    if (!ValidAdapterString(identity.candidate_id, 256) ||
        !ValidAdapterString(identity.upstream_pin, 256) ||
        !ValidAdapterString(identity.source_sha256, 64) ||
        !ValidAdapterString(identity.adapter_header_sha256, 64) ||
        !ValidAdapterString(identity.build_identity, 80))
    {
        return false;
    }
    canonical.append("{\"adapterAbi\":");
    AppendUnsigned(canonical, identity.adapter_abi);
    canonical.append(",\"adapterHeaderSha256\":");
    AppendJcsString(
        canonical,
        AdapterStringView(identity.adapter_header_sha256));
    canonical.append(",\"architecture\":");
    AppendUnsigned(canonical, identity.architecture);
    canonical.append(",\"buildIdentity\":");
    AppendJcsString(
        canonical,
        AdapterStringView(identity.build_identity));
    canonical.append(",\"callingConvention\":");
    AppendUnsigned(canonical, identity.calling_convention);
    canonical.append(",\"candidateId\":");
    AppendJcsString(
        canonical,
        AdapterStringView(identity.candidate_id));
    canonical.append(",\"capabilities\":");
    AppendUnsigned(canonical, identity.capabilities);
    canonical.append(",\"pointerSize\":");
    AppendUnsigned(canonical, identity.pointer_size);
    canonical.append(",\"sourceSha256\":");
    AppendJcsString(
        canonical,
        AdapterStringView(identity.source_sha256));
    canonical.append(",\"upstreamPin\":");
    AppendJcsString(
        canonical,
        AdapterStringView(identity.upstream_pin));
    canonical.push_back('}');
    return true;
}

[[nodiscard]] bool QueryAdapterLayoutJcs(
    HMODULE module,
    const Arguments& arguments,
    std::string& adapterLayoutSha256)
{
    const FARPROC address = GetProcAddress(
        module,
        kRsTerminalGate0QueryLayoutJcsSymbol);
    if (address == nullptr)
    {
        return false;
    }
    static_assert(
        sizeof(address) ==
        sizeof(RsTerminalGate0QueryLayoutJcsFn));
    const auto query =
        std::bit_cast<RsTerminalGate0QueryLayoutJcsFn>(address);

    uint64_t ignoredRequired = 0;
    if (query(
            kRsTerminalGate0LayoutQueryVersion + 1,
            nullptr,
            0,
            &ignoredRequired) !=
        RS_TERMINAL_GATE0_UNSUPPORTED_ABI)
    {
        return false;
    }

    uint64_t required = 0;
    if (query(
            kRsTerminalGate0LayoutQueryVersion,
            nullptr,
            0,
            &required) != RS_TERMINAL_GATE0_SUCCESS ||
        required == 0 ||
        required > kMaximumLayoutJcsBytes ||
        required != arguments.probeLayoutJcs.size())
    {
        return false;
    }
    std::vector<uint8_t> bytes(static_cast<size_t>(required));
    uint64_t secondRequired = required;
    if (query(
            kRsTerminalGate0LayoutQueryVersion,
            bytes.data(),
            required,
            &secondRequired) != RS_TERMINAL_GATE0_SUCCESS ||
        secondRequired != required)
    {
        return false;
    }
    const std::string_view adapterLayout(
        reinterpret_cast<const char*>(bytes.data()),
        bytes.size());
    return adapterLayout == arguments.probeLayoutJcs &&
        Sha256Hex(adapterLayout, adapterLayoutSha256) &&
        adapterLayoutSha256 == arguments.probeLayoutSha256;
}

[[nodiscard]] TraceResult RunTrace(
    const Arguments& arguments,
    const CorpusCase& corpusCase,
    uint64_t denialOrdinal)
{
    TraceResult trace;
    constexpr auto limits = MakeFuzzResourceLimits();
    ResourceLedger ledger(limits, denialOrdinal);
    const RsTerminalGate0ResourceCallbacksV1 callbacks = {
        .size = sizeof(RsTerminalGate0ResourceCallbacksV1),
        .userdata = &ledger,
        .reserve = ReserveResource,
        .release = ReleaseResource,
    };
    RsTerminalGate0Result createResult =
        RS_TERMINAL_GATE0_INTERNAL_ERROR;
    RsTerminalGate0Result stateProbeResult =
        RS_TERMINAL_GATE0_INTERNAL_ERROR;
    RsTerminalGate0Result stateReadResult =
        RS_TERMINAL_GATE0_INTERNAL_ERROR;
    RsTerminalGate0Result destroyResult =
        RS_TERMINAL_GATE0_INTERNAL_ERROR;
    std::vector<RsTerminalGate0Result> writeResults;
    bool loadSucceeded = false;
    bool layoutValid = false;
    bool identityValid = false;
    bool stateBounded = true;
    bool quiet = false;
    bool modulePathKnown = false;
    std::array<wchar_t, 32768> loadedModulePath{};

    wil::unique_hmodule module(LoadLibraryExW(
        arguments.adapterDll.c_str(),
        nullptr,
        LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR |
            LOAD_LIBRARY_SEARCH_SYSTEM32));
    if (module)
    {
        loadSucceeded = true;
        const DWORD pathLength = GetModuleFileNameW(
            module.get(),
            loadedModulePath.data(),
            static_cast<DWORD>(loadedModulePath.size()));
        modulePathKnown =
            pathLength != 0 &&
            pathLength < loadedModulePath.size();

        layoutValid = QueryAdapterLayoutJcs(
            module.get(),
            arguments,
            trace.adapterLayoutSha256);
        if (layoutValid)
        {
            const FARPROC address = GetProcAddress(
                module.get(),
                "rs_terminal_gate0_query_adapter");
            if (address != nullptr)
            {
                static_assert(
                    sizeof(address) ==
                    sizeof(RsTerminalGate0QueryAdapterFn));
                const auto query =
                    std::bit_cast<RsTerminalGate0QueryAdapterFn>(
                        address);
                RsTerminalGate0AdapterV1 adapter{};
                adapter.size = sizeof(adapter);
                adapter.identity.size = sizeof(adapter.identity);
                const RsTerminalGate0Result queryResult = query(
                    RS_TERMINAL_GATE0_QUERY_SYMBOL_VERSION,
                    RS_TERMINAL_GATE0_ADAPTER_ABI_V1,
                    &adapter);
                std::string identityCanonical;
                identityValid =
                    queryResult == RS_TERMINAL_GATE0_SUCCESS &&
                    adapter.size == sizeof(adapter) &&
                    adapter.identity.size ==
                        sizeof(adapter.identity) &&
                    adapter.identity.adapter_abi ==
                        RS_TERMINAL_GATE0_ADAPTER_ABI_V1 &&
                    adapter.identity.architecture ==
                        kExpectedArchitecture &&
                    adapter.identity.calling_convention ==
                        RS_TERMINAL_GATE0_CALL_CDECL &&
                    adapter.identity.pointer_size == sizeof(void*) &&
                    adapter.identity.reserved == 0 &&
                    AdapterStringView(
                        adapter.identity.candidate_id) ==
                        arguments.expectedCandidateId &&
                    AdapterStringView(
                        adapter.identity.upstream_pin) ==
                        arguments.expectedUpstreamPin &&
                    AdapterStringView(
                        adapter.identity.source_sha256) ==
                        arguments.expectedSourceSha256 &&
                    AdapterStringView(
                        adapter.identity.adapter_header_sha256) ==
                        arguments.expectedAdapterHeaderSha256 &&
                    AdapterStringView(
                        adapter.identity.build_identity) ==
                        arguments.expectedBuildIdentity &&
                    adapter.identity.capabilities ==
                        arguments.expectedCapabilities &&
                    adapter.session_create != nullptr &&
                    adapter.session_write != nullptr &&
                    adapter.session_action != nullptr &&
                    adapter.session_state_jcs != nullptr &&
                    adapter.session_destroy != nullptr &&
                    AppendAdapterIdentity(
                        adapter.identity,
                        identityCanonical) &&
                    Sha256Hex(
                        identityCanonical,
                        trace.adapterIdentitySha256);
                if (identityValid)
                {
                    RsTerminalGate0Session session = nullptr;
                    const RsTerminalGate0SessionConfigV1 config = {
                        .size =
                            sizeof(
                                RsTerminalGate0SessionConfigV1),
                        .columns = 120,
                        .rows = 40,
                        .cell_width_px = 8,
                        .cell_height_px = 16,
                        .fixture_id =
                            AdapterString(kFuzzFixtureId),
                        .schedule_id =
                            AdapterString(kFuzzScheduleId),
                        .limits = &limits,
                        .resource_callbacks = &callbacks,
                    };
                createResult =
                    adapter.session_create(&config, &session);
                if (createResult ==
                        RS_TERMINAL_GATE0_SUCCESS &&
                    session != nullptr)
                {
                    size_t offset = 0;
                    writeResults.reserve(
                        corpusCase.chunks.size());
                    for (const size_t chunk :
                         corpusCase.chunks)
                    {
                        const uint8_t* data =
                            chunk == 0
                            ? nullptr
                            : corpusCase.bytes.data() +
                                offset;
                        const RsTerminalGate0Result result =
                            adapter.session_write(
                                session,
                                data,
                                chunk);
                        writeResults.push_back(result);
                        offset += chunk;
                        if (!IsSafeWriteResult(
                                result,
                                denialOrdinal != 0))
                        {
                            break;
                        }
                        if (result !=
                            RS_TERMINAL_GATE0_SUCCESS)
                        {
                            break;
                        }
                    }
                    if (offset > corpusCase.bytes.size())
                    {
                        stateBounded = false;
                    }

                    size_t required = 0;
                    stateProbeResult =
                        adapter.session_state_jcs(
                            session,
                            nullptr,
                            0,
                            &required);
                    if (required > kMaximumStateBytes)
                    {
                        stateBounded = false;
                    }
                    else if (required == 0)
                    {
                        std::string emptyHash;
                        if (Sha256Hex(
                                std::string_view{},
                                emptyHash))
                        {
                            trace.stateSha256 =
                                std::move(emptyHash);
                        }
                        stateReadResult =
                            stateProbeResult;
                    }
                    else
                    {
                        std::vector<uint8_t> state(required);
                        size_t secondRequired = required;
                        stateReadResult =
                            adapter.session_state_jcs(
                                session,
                                state.data(),
                                state.size(),
                                &secondRequired);
                        if (stateReadResult ==
                                RS_TERMINAL_GATE0_SUCCESS &&
                            secondRequired == required)
                        {
                            (void)Sha256Hex(
                                state,
                                trace.stateSha256);
                        }
                    }
                    destroyResult =
                        adapter.session_destroy(session);
                }
            }
        }
    }
    }

    ledger.Close();
    quiet = ledger.IsQuiet();
    trace.reserveCalls = ledger.ReserveCalls();
    trace.denialObserved = ledger.DenialObserved();
    trace.resourceCanonical = ledger.CanonicalTrace();
    if (trace.stateSha256.empty())
    {
        (void)Sha256Hex(
            std::string_view{},
            trace.stateSha256);
    }

    module.reset();
    trace.unloaded =
        loadSucceeded &&
        modulePathKnown &&
        GetModuleHandleW(loadedModulePath.data()) == nullptr;

    trace.resultCanonical.append("{\"create\":");
    AppendUnsigned(trace.resultCanonical, createResult);
    trace.resultCanonical.append(",\"destroy\":");
    AppendUnsigned(trace.resultCanonical, destroyResult);
    trace.resultCanonical.append(",\"identity\":");
    trace.resultCanonical.append(
        identityValid ? "true" : "false");
    trace.resultCanonical.append(",\"layout\":");
    trace.resultCanonical.append(
        layoutValid ? "true" : "false");
    trace.resultCanonical.append(",\"load\":");
    trace.resultCanonical.append(
        loadSucceeded ? "true" : "false");
    trace.resultCanonical.append(",\"stateProbe\":");
    AppendUnsigned(trace.resultCanonical, stateProbeResult);
    trace.resultCanonical.append(",\"stateRead\":");
    AppendUnsigned(trace.resultCanonical, stateReadResult);
    trace.resultCanonical.append(",\"unloaded\":");
    trace.resultCanonical.append(
        trace.unloaded ? "true" : "false");
    trace.resultCanonical.append(",\"writes\":[");
    for (size_t index = 0;
         index < writeResults.size();
         ++index)
    {
        if (index != 0)
        {
            trace.resultCanonical.push_back(',');
        }
        AppendUnsigned(
            trace.resultCanonical,
            writeResults[index]);
    }
    trace.resultCanonical.append("]}");

    const bool baselineCreate =
        denialOrdinal == 0 &&
        createResult == RS_TERMINAL_GATE0_SUCCESS;
    const bool deniedCreate =
        denialOrdinal != 0 &&
        (createResult ==
             RS_TERMINAL_GATE0_OUT_OF_MEMORY ||
         createResult ==
             RS_TERMINAL_GATE0_LIMIT_REJECTED);
    const bool sessionCreated =
        createResult == RS_TERMINAL_GATE0_SUCCESS;
    const bool writesSafe = std::ranges::all_of(
        writeResults,
        [&](RsTerminalGate0Result result) {
            return IsKnownResult(result) &&
                IsSafeWriteResult(
                    result,
                    denialOrdinal != 0);
        });
    trace.passed =
        layoutValid &&
        identityValid &&
        stateBounded &&
        (baselineCreate || deniedCreate || sessionCreated) &&
        (!sessionCreated ||
         destroyResult == RS_TERMINAL_GATE0_SUCCESS) &&
        writesSafe &&
        IsKnownResult(stateProbeResult) &&
        IsKnownResult(stateReadResult) &&
        quiet &&
        trace.unloaded;
    return trace;
}

[[nodiscard]] bool SameTrace(
    const TraceResult& left,
    const TraceResult& right) noexcept
{
    return left.adapterIdentitySha256 ==
            right.adapterIdentitySha256 &&
        left.adapterLayoutSha256 ==
            right.adapterLayoutSha256 &&
        left.resourceCanonical == right.resourceCanonical &&
        left.resultCanonical == right.resultCanonical &&
        left.stateSha256 == right.stateSha256 &&
        left.reserveCalls == right.reserveCalls &&
        left.denialObserved == right.denialObserved &&
        left.unloaded == right.unloaded &&
        left.passed == right.passed;
}

struct FuzzSummary
{
    std::string adapterIdentitySha256;
    std::string adapterLayoutSha256;
    std::string resourceTraceSha256;
    std::string resultTraceSha256;
    std::string stateTraceSha256;
    size_t caseCount = 0;
    uint64_t denialSweepCount = 0;
    uint64_t runCount = 0;
    uint64_t unloadCount = 0;
    bool passed = true;
};

void AppendDigestRecord(
    std::string& resultRecords,
    std::string& stateRecords,
    std::string& resourceRecords,
    const TraceResult& trace,
    bool& first)
{
    std::string resultHash;
    std::string resourceHash;
    (void)Sha256Hex(
        trace.resultCanonical,
        resultHash);
    (void)Sha256Hex(
        trace.resourceCanonical,
        resourceHash);
    if (!first)
    {
        resultRecords.push_back(',');
        stateRecords.push_back(',');
        resourceRecords.push_back(',');
    }
    first = false;
    AppendJcsString(resultRecords, resultHash);
    AppendJcsString(stateRecords, trace.stateSha256);
    AppendJcsString(resourceRecords, resourceHash);
}

[[nodiscard]] FuzzSummary RunFuzz(
    const Arguments& arguments)
{
    FuzzSummary summary;
    const std::vector<CorpusCase> corpus =
        BuildCorpus(arguments);
    summary.caseCount = corpus.size();
    std::string resultRecords("[");
    std::string stateRecords("[");
    std::string resourceRecords("[");
    bool firstRecord = true;
    std::array<std::string, 3> fixedStateHashes{};

    for (size_t caseIndex = 0;
         caseIndex < corpus.size();
         ++caseIndex)
    {
        const TraceResult first =
            RunTrace(arguments, corpus[caseIndex], 0);
        const TraceResult second =
            RunTrace(arguments, corpus[caseIndex], 0);
        summary.runCount += 2;
        summary.unloadCount +=
            static_cast<uint64_t>(first.unloaded) +
            static_cast<uint64_t>(second.unloaded);
        summary.passed =
            summary.passed &&
            first.passed &&
            second.passed &&
            SameTrace(first, second);
        if (summary.adapterIdentitySha256.empty())
        {
            summary.adapterIdentitySha256 =
                first.adapterIdentitySha256;
            summary.adapterLayoutSha256 =
                first.adapterLayoutSha256;
        }
        else
        {
            summary.passed =
                summary.passed &&
                summary.adapterIdentitySha256 ==
                    first.adapterIdentitySha256 &&
                summary.adapterLayoutSha256 ==
                    first.adapterLayoutSha256;
        }
        AppendDigestRecord(
            resultRecords,
            stateRecords,
            resourceRecords,
            first,
            firstRecord);
        if (caseIndex >= 1 && caseIndex <= 3)
        {
            fixedStateHashes[caseIndex - 1] =
                first.stateSha256;
        }
    }
    summary.passed =
        summary.passed &&
        fixedStateHashes[0] == fixedStateHashes[1];

    const CorpusCase denialCase = BuildDenialCase(corpus);
    const TraceResult denialBaseline =
        RunTrace(arguments, denialCase, 0);
    const TraceResult denialBaselineRepeat =
        RunTrace(arguments, denialCase, 0);
    summary.runCount += 2;
    summary.unloadCount +=
        static_cast<uint64_t>(denialBaseline.unloaded) +
        static_cast<uint64_t>(
            denialBaselineRepeat.unloaded);
    summary.passed =
        summary.passed &&
        denialBaseline.passed &&
        denialBaselineRepeat.passed &&
        SameTrace(denialBaseline, denialBaselineRepeat) &&
        denialBaseline.reserveCalls != 0 &&
        denialBaseline.reserveCalls <=
            kMaximumDenialOrdinals;
    AppendDigestRecord(
        resultRecords,
        stateRecords,
        resourceRecords,
        denialBaseline,
        firstRecord);

    if (denialBaseline.reserveCalls <=
        kMaximumDenialOrdinals)
    {
        summary.denialSweepCount =
            denialBaseline.reserveCalls;
        for (uint64_t ordinal = 1;
             ordinal <= summary.denialSweepCount;
             ++ordinal)
        {
            const TraceResult first =
                RunTrace(arguments, denialCase, ordinal);
            const TraceResult second =
                RunTrace(arguments, denialCase, ordinal);
            summary.runCount += 2;
            summary.unloadCount +=
                static_cast<uint64_t>(first.unloaded) +
                static_cast<uint64_t>(second.unloaded);
            summary.passed =
                summary.passed &&
                first.passed &&
                second.passed &&
                first.denialObserved &&
                second.denialObserved &&
                SameTrace(first, second);
            AppendDigestRecord(
                resultRecords,
                stateRecords,
                resourceRecords,
                first,
                firstRecord);
        }
    }

    resultRecords.push_back(']');
    stateRecords.push_back(']');
    resourceRecords.push_back(']');
    summary.passed =
        summary.passed &&
        summary.adapterLayoutSha256 ==
            arguments.probeLayoutSha256 &&
        summary.unloadCount == summary.runCount &&
        Sha256Hex(
            resultRecords,
            summary.resultTraceSha256) &&
        Sha256Hex(
            stateRecords,
            summary.stateTraceSha256) &&
        Sha256Hex(
            resourceRecords,
            summary.resourceTraceSha256);
    return summary;
}

struct C7SessionResult
{
    std::string stateJcs;
    uint64_t actionFlags = 0;
    bool quiet = false;
    bool passed = false;
};

struct C7ScheduleState
{
    std::string_view scheduleId;
    std::string stateJcs;
};

struct C7Summary
{
    std::array<C7ScheduleState, 3> scheduleStates = {
        C7ScheduleState{"all-at-once", {}},
        C7ScheduleState{"bytewise", {}},
        C7ScheduleState{"terminator-split", {}},
    };
    std::string adapterIdentitySha256;
    std::string adapterLayoutSha256;
    std::string orderingStateJcs;
    uint64_t orderingActionFlags = 0;
    uint64_t quietCount = 0;
    bool unloaded = false;
    bool passed = false;
};

[[nodiscard]] std::vector<size_t> C7ChunkEnds(
    std::span<const uint8_t> bytes,
    std::string_view scheduleId)
{
    std::vector<size_t> ends;
    if (scheduleId == "all-at-once")
    {
        ends.push_back(bytes.size());
    }
    else if (scheduleId == "bytewise")
    {
        ends.reserve(bytes.size());
        for (size_t index = 1; index <= bytes.size(); ++index)
        {
            ends.push_back(index);
        }
    }
    else if (scheduleId == "terminator-split")
    {
        for (size_t index = 0; index + 1 < bytes.size(); ++index)
        {
            if (bytes[index] == 0x1B &&
                bytes[index + 1] == static_cast<uint8_t>('\\'))
            {
                ends.push_back(index + 1);
            }
        }
        if (ends.empty() || ends.back() != bytes.size())
        {
            ends.push_back(bytes.size());
        }
    }
    return ends;
}

[[nodiscard]] bool ReadC7StateJcs(
    const RsTerminalGate0AdapterV1& adapter,
    RsTerminalGate0Session session,
    std::string& state)
{
    size_t required = 0;
    if (adapter.session_state_jcs(
            session,
            nullptr,
            0,
            &required) != RS_TERMINAL_GATE0_OUT_OF_SPACE ||
        required == 0 ||
        required > kMaximumC7StateBytes)
    {
        return false;
    }
    state.assign(required, '\0');
    size_t written = 0;
    return adapter.session_state_jcs(
               session,
               reinterpret_cast<uint8_t*>(state.data()),
               state.size(),
               &written) == RS_TERMINAL_GATE0_SUCCESS &&
        written == required;
}

[[nodiscard]] C7SessionResult RunC7Session(
    const RsTerminalGate0AdapterV1& adapter,
    std::string_view fixtureId,
    std::string_view scheduleId,
    std::span<const uint8_t> input,
    std::span<const size_t> chunkEnds,
    std::string_view actionId,
    std::string_view actionArguments)
{
    C7SessionResult result;
    constexpr auto limits = MakeFuzzResourceLimits();
    ResourceLedger ledger(limits, 0);
    const RsTerminalGate0ResourceCallbacksV1 callbacks = {
        .size = sizeof(RsTerminalGate0ResourceCallbacksV1),
        .userdata = &ledger,
        .reserve = ReserveResource,
        .release = ReleaseResource,
    };
    const RsTerminalGate0SessionConfigV1 config = {
        .size = sizeof(RsTerminalGate0SessionConfigV1),
        .columns = 120,
        .rows = 40,
        .cell_width_px = 8,
        .cell_height_px = 16,
        .fixture_id = AdapterString(fixtureId),
        .schedule_id = AdapterString(scheduleId),
        .limits = &limits,
        .resource_callbacks = &callbacks,
    };

    RsTerminalGate0Session session = nullptr;
    bool createSucceeded =
        adapter.session_create(&config, &session) ==
            RS_TERMINAL_GATE0_SUCCESS &&
        session != nullptr;
    bool writeSucceeded = createSucceeded &&
        !chunkEnds.empty();
    size_t offset = 0;
    if (writeSucceeded)
    {
        for (const size_t end : chunkEnds)
        {
            if (end <= offset ||
                end > input.size() ||
                adapter.session_write(
                    session,
                    input.data() + offset,
                    end - offset) !=
                    RS_TERMINAL_GATE0_SUCCESS)
            {
                writeSucceeded = false;
                break;
            }
            offset = end;
        }
        writeSucceeded =
            writeSucceeded &&
            offset == input.size();
    }

    bool actionSucceeded = false;
    if (writeSucceeded)
    {
        const RsTerminalGate0ActionV1 action = {
            .size = sizeof(RsTerminalGate0ActionV1),
            .action_id = AdapterString(actionId),
            .arguments_jcs = AdapterString(actionArguments),
        };
        RsTerminalGate0ActionObservationV1 observation{};
        observation.size = sizeof(observation);
        actionSucceeded =
            adapter.session_action(
                session,
                &action,
                &observation) ==
                RS_TERMINAL_GATE0_SUCCESS &&
            observation.size == sizeof(observation);
        result.actionFlags = observation.flags;
    }
    const bool stateRead =
        actionSucceeded &&
        ReadC7StateJcs(
            adapter,
            session,
            result.stateJcs);
    bool destroySucceeded = false;
    if (session != nullptr)
    {
        destroySucceeded =
            adapter.session_destroy(session) ==
            RS_TERMINAL_GATE0_SUCCESS;
    }
    ledger.Close();
    result.quiet =
        destroySucceeded &&
        ledger.IsQuiet();
    result.passed =
        createSucceeded &&
        writeSucceeded &&
        actionSucceeded &&
        stateRead &&
        result.quiet;
    return result;
}

[[nodiscard]] bool LoadC7Adapter(
    HMODULE module,
    const Arguments& arguments,
    RsTerminalGate0AdapterV1& adapter,
    std::string& adapterIdentitySha256,
    std::string& adapterLayoutSha256)
{
    if (!QueryAdapterLayoutJcs(
            module,
            arguments,
            adapterLayoutSha256))
    {
        return false;
    }
    const FARPROC address = GetProcAddress(
        module,
        "rs_terminal_gate0_query_adapter");
    if (address == nullptr)
    {
        return false;
    }
    static_assert(
        sizeof(address) ==
        sizeof(RsTerminalGate0QueryAdapterFn));
    const auto query =
        std::bit_cast<RsTerminalGate0QueryAdapterFn>(address);
    adapter = {};
    adapter.size = sizeof(adapter);
    adapter.identity.size = sizeof(adapter.identity);
    if (query(
            RS_TERMINAL_GATE0_QUERY_SYMBOL_VERSION,
            RS_TERMINAL_GATE0_ADAPTER_ABI_V1,
            &adapter) != RS_TERMINAL_GATE0_SUCCESS)
    {
        return false;
    }
    std::string identityCanonical;
    return adapter.size == sizeof(adapter) &&
        adapter.identity.size == sizeof(adapter.identity) &&
        adapter.identity.adapter_abi ==
            RS_TERMINAL_GATE0_ADAPTER_ABI_V1 &&
        adapter.identity.architecture ==
            kExpectedArchitecture &&
        adapter.identity.calling_convention ==
            RS_TERMINAL_GATE0_CALL_CDECL &&
        adapter.identity.pointer_size == sizeof(void*) &&
        adapter.identity.reserved == 0 &&
        AdapterStringView(adapter.identity.candidate_id) ==
            arguments.expectedCandidateId &&
        AdapterStringView(adapter.identity.upstream_pin) ==
            arguments.expectedUpstreamPin &&
        AdapterStringView(adapter.identity.source_sha256) ==
            arguments.expectedSourceSha256 &&
        AdapterStringView(
            adapter.identity.adapter_header_sha256) ==
            arguments.expectedAdapterHeaderSha256 &&
        AdapterStringView(adapter.identity.build_identity) ==
            arguments.expectedBuildIdentity &&
        adapter.identity.capabilities ==
            arguments.expectedCapabilities &&
        adapter.session_create != nullptr &&
        adapter.session_write != nullptr &&
        adapter.session_action != nullptr &&
        adapter.session_state_jcs != nullptr &&
        adapter.session_destroy != nullptr &&
        AppendAdapterIdentity(
            adapter.identity,
            identityCanonical) &&
        Sha256Hex(
            identityCanonical,
            adapterIdentitySha256);
}

[[nodiscard]] C7Summary RunC7(
    const Arguments& arguments)
{
    C7Summary summary;
    std::array<wchar_t, 32768> loadedModulePath{};
    bool modulePathKnown = false;
    wil::unique_hmodule module(LoadLibraryExW(
        arguments.adapterDll.c_str(),
        nullptr,
        LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR |
            LOAD_LIBRARY_SEARCH_SYSTEM32));
    RsTerminalGate0AdapterV1 adapter{};
    bool identityValid = false;
    if (module)
    {
        const DWORD pathLength = GetModuleFileNameW(
            module.get(),
            loadedModulePath.data(),
            static_cast<DWORD>(loadedModulePath.size()));
        modulePathKnown =
            pathLength != 0 &&
            pathLength < loadedModulePath.size();
        identityValid = LoadC7Adapter(
            module.get(),
            arguments,
            adapter,
            summary.adapterIdentitySha256,
            summary.adapterLayoutSha256);
    }

    bool schedulesPassed = identityValid;
    if (identityValid)
    {
        for (C7ScheduleState& schedule :
             summary.scheduleStates)
        {
            const std::vector<size_t> chunks =
                C7ChunkEnds(
                    arguments.c7Stream,
                    schedule.scheduleId);
            C7SessionResult result = RunC7Session(
                adapter,
                kC7FixtureId,
                schedule.scheduleId,
                arguments.c7Stream,
                chunks,
                kC7ActionId,
                kC7ActionArguments);
            schedule.stateJcs =
                std::move(result.stateJcs);
            summary.quietCount +=
                static_cast<uint64_t>(result.quiet);
            schedulesPassed =
                schedulesPassed &&
                result.passed;
        }
    }

    bool orderingPassed = false;
    if (identityValid)
    {
        const auto& bytes =
            FixtureContract::k_write_response_order_input;
        std::vector<uint8_t> orderingInput;
        const std::string_view hex = bytes[0].hex;
        orderingInput.reserve(hex.size() / 2);
        for (size_t index = 0; index < hex.size(); index += 2)
        {
            const auto digit = [](char character) -> uint8_t {
                return character <= '9'
                    ? static_cast<uint8_t>(character - '0')
                    : static_cast<uint8_t>(
                        character - 'a' + 10);
            };
            orderingInput.push_back(
                static_cast<uint8_t>(
                    (digit(hex[index]) << 4) |
                    digit(hex[index + 1])));
        }
        std::vector<size_t> chunkEnds;
        for (const uint64_t cut :
             FixtureContract::
                 k_write_response_order_terminator_split_cuts)
        {
            chunkEnds.push_back(
                static_cast<size_t>(cut));
        }
        chunkEnds.push_back(orderingInput.size());
        C7SessionResult ordering = RunC7Session(
            adapter,
            "write-response-order",
            "terminator-split",
            orderingInput,
            chunkEnds,
            "admit-external-events",
            "{\"order\":\"input,resize,side-channel\"}");
        constexpr uint64_t requiredOrderingFlags =
            RS_TERMINAL_GATE0_ACTION_OBSERVED |
            RS_TERMINAL_GATE0_ACTION_CHECKED_ARITHMETIC |
            RS_TERMINAL_GATE0_ACTION_ORDER_PRESERVED;
        summary.orderingStateJcs =
            std::move(ordering.stateJcs);
        summary.orderingActionFlags =
            ordering.actionFlags;
        summary.quietCount +=
            static_cast<uint64_t>(ordering.quiet);
        orderingPassed =
            ordering.passed &&
            ordering.actionFlags ==
                requiredOrderingFlags;
    }

    module.reset();
    summary.unloaded =
        modulePathKnown &&
        GetModuleHandleW(loadedModulePath.data()) ==
            nullptr;
    summary.passed =
        identityValid &&
        schedulesPassed &&
        orderingPassed &&
        summary.adapterLayoutSha256 ==
            arguments.probeLayoutSha256 &&
        summary.quietCount == 4 &&
        summary.unloaded;
    return summary;
}

void EmitC7(
    const Arguments& arguments,
    const C7Summary& summary)
{
    std::string output;
    output.reserve(32768);
    output.append("{\"adapterHeaderSha256\":");
    AppendJcsString(
        output,
        arguments.expectedAdapterHeaderSha256);
    output.append(",\"adapterIdentitySha256\":");
    AppendJcsString(
        output,
        summary.adapterIdentitySha256);
    output.append(",\"adapterLayoutSha256\":");
    AppendJcsString(
        output,
        summary.adapterLayoutSha256);
    output.append(",\"architecture\":");
    AppendJcsString(
        output,
        arguments.expectedArchitecture);
    output.append(",\"buildIdentity\":");
    AppendJcsString(
        output,
        arguments.expectedBuildIdentity);
    output.append(",\"candidateId\":");
    AppendJcsString(
        output,
        arguments.expectedCandidateId);
    output.append(",\"capabilities\":");
    AppendUnsigned(
        output,
        arguments.expectedCapabilities);
    std::string fixedStreamSha256;
    (void)Sha256Hex(
        arguments.c7Stream,
        fixedStreamSha256);
    output.append(",\"fixedStreamSha256\":");
    AppendJcsString(
        output,
        fixedStreamSha256);
    output.append(
        ",\"formatId\":\"red-salamander-terminal-engine-gate0-c7\","
        "\"orderingActionFlags\":[");
    constexpr std::array<std::pair<uint64_t, std::string_view>, 3>
        orderingFlags = {{
            {
                RS_TERMINAL_GATE0_ACTION_OBSERVED,
                "action-observed",
            },
            {
                RS_TERMINAL_GATE0_ACTION_CHECKED_ARITHMETIC,
                "checked-arithmetic",
            },
            {
                RS_TERMINAL_GATE0_ACTION_ORDER_PRESERVED,
                "order-preserved",
            },
        }};
    bool firstFlag = true;
    for (const auto& [flag, name] : orderingFlags)
    {
        if ((summary.orderingActionFlags & flag) == 0)
        {
            continue;
        }
        if (!firstFlag)
        {
            output.push_back(',');
        }
        firstFlag = false;
        AppendJcsString(output, name);
    }
    output.append("],\"orderingStateJcs\":");
    AppendJcsString(
        output,
        summary.orderingStateJcs);
    output.append(",\"quietCount\":");
    AppendUnsigned(output, summary.quietCount);
    output.append(",\"scheduleStates\":[");
    for (size_t index = 0;
         index < summary.scheduleStates.size();
         ++index)
    {
        if (index != 0)
        {
            output.push_back(',');
        }
        output.append("{\"scheduleId\":");
        AppendJcsString(
            output,
            summary.scheduleStates[index].scheduleId);
        output.append(",\"stateJcs\":");
        AppendJcsString(
            output,
            summary.scheduleStates[index].stateJcs);
        output.push_back('}');
    }
    output.append("],\"sessionCount\":4,\"sourceSha256\":");
    AppendJcsString(
        output,
        arguments.expectedSourceSha256);
    output.append(",\"status\":");
    AppendJcsString(
        output,
        summary.passed ? "pass" : "fail");
    output.append(",\"unloadCount\":");
    AppendUnsigned(
        output,
        summary.unloaded ? 1 : 0);
    output.append(",\"upstreamPin\":");
    AppendJcsString(
        output,
        arguments.expectedUpstreamPin);
    output.append(",\"version\":1}\n");
    std::fwrite(
        output.data(),
        1,
        output.size(),
        stdout);
}

void EmitLayout(
    std::string_view layoutJcs,
    std::string_view layoutSha256)
{
    std::string output;
    output.reserve(layoutJcs.size() + 512);
    output.append(
        "{\"formatId\":\"red-salamander-terminal-engine-gate0-layout\","
        "\"layoutIdentity\":");
    const std::string identity =
        std::string("layout-v1-") +
        std::string(layoutSha256);
    AppendJcsString(output, identity);
    output.append(",\"layoutJcs\":");
    output.append(layoutJcs);
    output.append(",\"layoutSha256\":");
    AppendJcsString(output, layoutSha256);
    output.append(
        ",\"status\":\"pass\",\"version\":1}\n");
    std::fwrite(
        output.data(),
        1,
        output.size(),
        stdout);
}

void EmitFuzz(
    const Arguments& arguments,
    const FuzzSummary& summary)
{
    std::string output;
    output.reserve(1024);
    output.append("{\"adapterHeaderSha256\":");
    AppendJcsString(
        output,
        arguments.expectedAdapterHeaderSha256);
    output.append(",\"adapterIdentitySha256\":");
    AppendJcsString(
        output,
        summary.adapterIdentitySha256);
    output.append(",\"adapterLayoutSha256\":");
    AppendJcsString(
        output,
        summary.adapterLayoutSha256);
    output.append(",\"architecture\":");
    AppendJcsString(
        output,
        arguments.expectedArchitecture);
    output.append(",\"buildIdentity\":");
    AppendJcsString(
        output,
        arguments.expectedBuildIdentity);
    output.append(",\"candidateId\":");
    AppendJcsString(
        output,
        arguments.expectedCandidateId);
    output.append(",\"capabilities\":");
    AppendUnsigned(
        output,
        arguments.expectedCapabilities);
    output.append(",\"caseCount\":");
    AppendUnsigned(output, summary.caseCount);
    output.append(",\"denialSweepCount\":");
    AppendUnsigned(output, summary.denialSweepCount);
    output.append(
        ",\"formatId\":\"red-salamander-terminal-engine-gate0-fuzz\","
        "\"layoutIdentity\":");
    AppendJcsString(
        output,
        arguments.expectedLayoutIdentity);
    output.append(",\"repeatCount\":2,\"resourceTraceSha256\":");
    AppendJcsString(
        output,
        summary.resourceTraceSha256);
    output.append(",\"resultTraceSha256\":");
    AppendJcsString(
        output,
        summary.resultTraceSha256);
    output.append(",\"runCount\":");
    AppendUnsigned(output, summary.runCount);
    output.append(",\"seed\":\"");
    AppendUnsigned(output, arguments.seed);
    output.push_back('"');
    output.append(",\"sourceSha256\":");
    AppendJcsString(
        output,
        arguments.expectedSourceSha256);
    output.append(",\"stateTraceSha256\":");
    AppendJcsString(
        output,
        summary.stateTraceSha256);
    output.append(",\"status\":");
    AppendJcsString(
        output,
        summary.passed ? "pass" : "fail");
    output.append(",\"unloadCount\":");
    AppendUnsigned(output, summary.unloadCount);
    output.append(",\"upstreamPin\":");
    AppendJcsString(
        output,
        arguments.expectedUpstreamPin);
    output.append(",\"version\":1}\n");
    std::fwrite(
        output.data(),
        1,
        output.size(),
        stdout);
}
} // namespace

int wmain(int argc, wchar_t** argv)
{
    Arguments arguments;
    if (!ParseArguments(argc, argv, arguments))
    {
        std::fputs(
            "Invalid arguments for TerminalEngineGate0Probe.\n",
            stderr);
        return 2;
    }

    const std::string layoutJcs = BuildLayoutJcs();
    std::string layoutSha256;
    if (!Sha256Hex(layoutJcs, layoutSha256))
    {
        std::fputs(
            "Failed to hash the Gate-0 ABI layout.\n",
            stderr);
        return 3;
    }
    const std::string layoutIdentity =
        std::string("layout-v1-") + layoutSha256;
    if (arguments.mode == Mode::Layout)
    {
        EmitLayout(layoutJcs, layoutSha256);
        return 0;
    }
    if (arguments.expectedLayoutIdentity != layoutIdentity)
    {
        std::fputs(
            "Expected Gate-0 ABI layout identity does not match this build.\n",
            stderr);
        return 4;
    }
    arguments.probeLayoutJcs = layoutJcs;
    arguments.probeLayoutSha256 = layoutSha256;

    if (arguments.mode == Mode::C7)
    {
        const C7Summary summary = RunC7(arguments);
        EmitC7(arguments, summary);
        return summary.passed ? 0 : 6;
    }
    const FuzzSummary summary = RunFuzz(arguments);
    EmitFuzz(arguments, summary);
    return summary.passed ? 0 : 5;
}

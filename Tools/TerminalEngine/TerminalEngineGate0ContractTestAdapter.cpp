// Harness mechanics test double only. This source intentionally reads the
// frozen expected values and MUST NOT be used as candidate or promotion
// evidence, shipped, or linked into the product.
#define RS_TERMINAL_GATE0_ADAPTER_EXPORTS
#include "TerminalEngineGate0Layout.h"

#include <TerminalEngineFixtures.v1.h>
#include <Windows.h>

#include <algorithm>
#include <array>
#include <cstring>
#include <memory>
#include <ranges>
#include <string_view>

namespace
{
namespace Fixtures = RedSalamander::TerminalEngine::Fixtures::V1;

#if defined(_M_X64)
constexpr RsTerminalGate0Architecture kArchitecture = RS_TERMINAL_GATE0_ARCH_X64;
#elif defined(_M_ARM64)
constexpr RsTerminalGate0Architecture kArchitecture = RS_TERMINAL_GATE0_ARCH_ARM64;
#else
#error Contract-test adapter supports only x64 and ARM64.
#endif

constexpr uint64_t kCapabilities = RS_TERMINAL_GATE0_CAP_CELLS_AND_STYLES | RS_TERMINAL_GATE0_CAP_SNAPSHOT_COPY | RS_TERMINAL_GATE0_CAP_INPUT_ENCODING |
                                   RS_TERMINAL_GATE0_CAP_EFFECT_ORDERING | RS_TERMINAL_GATE0_CAP_KITTY_STATIC | RS_TERMINAL_GATE0_CAP_RESOURCE_ACCOUNTING |
                                   RS_TERMINAL_GATE0_CAP_QUIET_DESTROY;

enum class TestMode
{
    Normal,
    StateMismatch,
    ActionUnsupported,
    LimitEvidenceMissing,
    WrongCapabilities,
};

[[nodiscard]] TestMode GetTestMode() noexcept
{
    constexpr wchar_t variableName[] = L"RS_TERMINAL_GATE0_CONTRACT_TEST_MODE";
    std::array<wchar_t, 64> value{};
    const DWORD length = GetEnvironmentVariableW(variableName, value.data(), static_cast<DWORD>(value.size()));
    if (length == 0 || length >= value.size())
    {
        return TestMode::Normal;
    }
    const std::wstring_view mode(value.data(), static_cast<size_t>(length));
    if (mode == L"state-mismatch")
    {
        return TestMode::StateMismatch;
    }
    if (mode == L"action-unsupported")
    {
        return TestMode::ActionUnsupported;
    }
    if (mode == L"limit-evidence-missing")
    {
        return TestMode::LimitEvidenceMissing;
    }
    if (mode == L"wrong-capabilities")
    {
        return TestMode::WrongCapabilities;
    }
    return TestMode::Normal;
}

[[nodiscard]] RsTerminalGate0String MakeString(std::string_view value) noexcept
{
    return {
        .ptr = reinterpret_cast<const uint8_t*>(value.data()),
        .len = value.size(),
    };
}

[[nodiscard]] std::string_view View(RsTerminalGate0String value) noexcept
{
    if (value.len != 0 && value.ptr == nullptr)
    {
        return {};
    }
    return {
        reinterpret_cast<const char*>(value.ptr),
        value.len,
    };
}

struct TestSession
{
    std::string_view fixtureId;
    const RsTerminalGate0ResourceCallbacksV1* callbacks = nullptr;
    bool reserved                                       = false;
};

[[nodiscard]] const Fixtures::Fixture* FindFixture(std::string_view fixtureId) noexcept
{
    const auto fixtures = Fixtures::GetFixtures();
    const auto found    = std::ranges::find(fixtures, fixtureId, &Fixtures::Fixture::fixtureId);
    return found == fixtures.end() ? nullptr : &*found;
}

[[nodiscard]] bool IsLimitAction(std::string_view actionId) noexcept
{
    constexpr std::string_view prefix = "probe-";
    return actionId.starts_with(prefix);
}

[[nodiscard]] bool IsOrderingAction(std::string_view actionId) noexcept
{
    return actionId == "encode-paste-enabled" || actionId == "encode-focus-gained-lost" || actionId == "encode-mouse-matrix" ||
           actionId == "admit-external-events";
}

RsTerminalGate0Result RS_TERMINAL_GATE0_CALL SessionCreate(const RsTerminalGate0SessionConfigV1* config, RsTerminalGate0Session* outSession) noexcept
{
    if (config == nullptr || outSession == nullptr || config->size != sizeof(*config) || config->limits == nullptr ||
        config->limits->size != sizeof(*config->limits) || config->resource_callbacks == nullptr ||
        config->resource_callbacks->size != sizeof(*config->resource_callbacks) || config->resource_callbacks->reserve == nullptr ||
        config->resource_callbacks->release == nullptr)
    {
        return RS_TERMINAL_GATE0_INVALID_ARGUMENT;
    }
    *outSession                      = nullptr;
    auto session                     = std::make_unique<TestSession>();
    session->fixtureId               = View(config->fixture_id);
    session->callbacks               = config->resource_callbacks;
    constexpr uint64_t baselineBytes = 4096;
    if (! session->callbacks->reserve(session->callbacks->userdata, RS_TERMINAL_GATE0_RESOURCE_TERMINAL_TEXT_CPU, baselineBytes, 0))
    {
        return RS_TERMINAL_GATE0_LIMIT_REJECTED;
    }
    session->reserved = true;
    *outSession       = session.release();
    return RS_TERMINAL_GATE0_SUCCESS;
}

RsTerminalGate0Result RS_TERMINAL_GATE0_CALL SessionWrite(RsTerminalGate0Session session, const uint8_t* data, size_t length) noexcept
{
    if (session == nullptr || (length != 0 && data == nullptr))
    {
        return RS_TERMINAL_GATE0_INVALID_ARGUMENT;
    }
    return RS_TERMINAL_GATE0_SUCCESS;
}

RsTerminalGate0Result RS_TERMINAL_GATE0_CALL SessionAction(RsTerminalGate0Session session,
                                                           const RsTerminalGate0ActionV1* action,
                                                           RsTerminalGate0ActionObservationV1* observation) noexcept
{
    if (session == nullptr || action == nullptr || action->size != sizeof(*action) || observation == nullptr || observation->size != sizeof(*observation))
    {
        return RS_TERMINAL_GATE0_INVALID_ARGUMENT;
    }
    const std::string_view actionId = View(action->action_id);
    if (GetTestMode() == TestMode::ActionUnsupported && ! actionId.starts_with("benchmark-"))
    {
        return RS_TERMINAL_GATE0_INTERNAL_ERROR;
    }
    uint64_t flags = RS_TERMINAL_GATE0_ACTION_OBSERVED | RS_TERMINAL_GATE0_ACTION_CHECKED_ARITHMETIC;
    if (IsLimitAction(actionId) && GetTestMode() != TestMode::LimitEvidenceMissing)
    {
        flags |= RS_TERMINAL_GATE0_ACTION_BOUNDARY_ACCEPTED | RS_TERMINAL_GATE0_ACTION_PLUS_ONE_REJECTED |
                 RS_TERMINAL_GATE0_ACTION_REJECTION_BEFORE_ALLOCATION | RS_TERMINAL_GATE0_ACTION_NO_PARTIAL_EFFECT;
    }
    if (IsOrderingAction(actionId))
    {
        flags |= RS_TERMINAL_GATE0_ACTION_ORDER_PRESERVED;
    }
    if (actionId == "teardown-sequence")
    {
        flags |= RS_TERMINAL_GATE0_ACTION_QUIET_POINT_REACHED;
    }
    observation->flags                                = flags;
    observation->reservation_calls_before_rejection   = 0;
    observation->allocation_attempts_before_rejection = 0;
    return RS_TERMINAL_GATE0_SUCCESS;
}

RsTerminalGate0Result RS_TERMINAL_GATE0_CALL SessionStateJcs(RsTerminalGate0Session rawSession, uint8_t* buffer, size_t capacity, size_t* outRequired) noexcept
{
    if (rawSession == nullptr || outRequired == nullptr)
    {
        return RS_TERMINAL_GATE0_INVALID_ARGUMENT;
    }
    const auto& session = *static_cast<TestSession*>(rawSession);
    const auto* fixture = FindFixture(session.fixtureId);
    if (fixture == nullptr)
    {
        return RS_TERMINAL_GATE0_INVALID_ARGUMENT;
    }
    const auto found = std::ranges::find(fixture->checks, std::string_view("state"), &Fixtures::Check::checkId);
    if (found == fixture->checks.end())
    {
        return RS_TERMINAL_GATE0_INTERNAL_ERROR;
    }
    constexpr std::string_view mismatch = "{\"contractTestCorruption\":true}";
    const std::string_view state        = GetTestMode() == TestMode::StateMismatch ? mismatch : found->expectedValue;
    *outRequired                        = state.size();
    if (buffer == nullptr || capacity < state.size())
    {
        return RS_TERMINAL_GATE0_OUT_OF_SPACE;
    }
    std::memcpy(buffer, state.data(), state.size());
    return RS_TERMINAL_GATE0_SUCCESS;
}

RsTerminalGate0Result RS_TERMINAL_GATE0_CALL SessionDestroy(RsTerminalGate0Session rawSession) noexcept
{
    if (rawSession == nullptr)
    {
        return RS_TERMINAL_GATE0_INVALID_ARGUMENT;
    }
    std::unique_ptr<TestSession> session(static_cast<TestSession*>(rawSession));
    if (session->reserved)
    {
        constexpr uint64_t baselineBytes = 4096;
        session->callbacks->release(session->callbacks->userdata, RS_TERMINAL_GATE0_RESOURCE_TERMINAL_TEXT_CPU, baselineBytes, 0);
        session->reserved = false;
    }
    return RS_TERMINAL_GATE0_SUCCESS;
}

} // namespace

extern "C" RS_TERMINAL_GATE0_API RsTerminalGate0Result RS_TERMINAL_GATE0_CALL rs_terminal_gate0_query_layout_jcs(uint32_t requestedVersion,
                                                                                                                 uint8_t* buffer,
                                                                                                                 uint64_t capacity,
                                                                                                                 uint64_t* outRequired) noexcept
{
    return RedSalamander::TerminalEngine::Gate0::Layout::QueryLayoutJcs(requestedVersion, buffer, capacity, outRequired);
}

extern "C" RS_TERMINAL_GATE0_API RsTerminalGate0Result RS_TERMINAL_GATE0_CALL rs_terminal_gate0_query_adapter(uint32_t querySymbolVersion,
                                                                                                              uint32_t requestedAdapterAbi,
                                                                                                              RsTerminalGate0AdapterV1* outAdapter)
{
    if (querySymbolVersion != RS_TERMINAL_GATE0_QUERY_SYMBOL_VERSION || requestedAdapterAbi != RS_TERMINAL_GATE0_ADAPTER_ABI_V1)
    {
        return RS_TERMINAL_GATE0_UNSUPPORTED_ABI;
    }
    if (outAdapter == nullptr || outAdapter->size != sizeof(*outAdapter) || outAdapter->identity.size != sizeof(outAdapter->identity))
    {
        return RS_TERMINAL_GATE0_INVALID_ARGUMENT;
    }
    constexpr std::string_view candidateId   = "contract-test-adapter";
    constexpr std::string_view upstreamPin   = "0000000000000000000000000000000000000000";
    constexpr std::string_view sourceSha256  = "0000000000000000000000000000000000000000000000000000000000000000";
    constexpr std::string_view headerSha256  = "5aed4f300a5a4aea090f50bbed743bd34824f18a9d2acc64d698a9b8561981bc";
    constexpr std::string_view buildIdentity = "build-v1-0000000000000000000000000000000000000000000000000000000000000000";
    const uint64_t capabilities = GetTestMode() == TestMode::WrongCapabilities ? kCapabilities & ~RS_TERMINAL_GATE0_CAP_KITTY_STATIC : kCapabilities;
    *outAdapter                 = {
        .size = sizeof(RsTerminalGate0AdapterV1),
        .identity =
            {
                .size                  = sizeof(RsTerminalGate0IdentityV1),
                .adapter_abi           = RS_TERMINAL_GATE0_ADAPTER_ABI_V1,
                .candidate_id          = MakeString(candidateId),
                .upstream_pin          = MakeString(upstreamPin),
                .source_sha256         = MakeString(sourceSha256),
                .adapter_header_sha256 = MakeString(headerSha256),
                .build_identity        = MakeString(buildIdentity),
                .architecture          = kArchitecture,
                .calling_convention    = RS_TERMINAL_GATE0_CALL_CDECL,
                .capabilities          = capabilities,
                .pointer_size          = sizeof(void*),
                .reserved              = 0,
            },
        .session_create    = SessionCreate,
        .session_write     = SessionWrite,
        .session_action    = SessionAction,
        .session_state_jcs = SessionStateJcs,
        .session_destroy   = SessionDestroy,
    };
    return RS_TERMINAL_GATE0_SUCCESS;
}

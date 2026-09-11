#pragma once

// Candidate-neutral Gate-0 ABI layout proof. Both the probe executable and
// the candidate adapter compile this serializer in their own translation
// units. The probe requires the adapter to export the exact JCS produced by
// QueryLayoutJcs, so an injected identity string cannot stand in for the
// adapter compiler's actual enum, function-pointer, and structure layout.

#include "TerminalEngineGate0Adapter.h"

#include <array>
#include <charconv>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <exception>
#include <new>
#include <stdexcept>
#include <string>
#include <string_view>

using RsTerminalGate0QueryLayoutJcsFn =
    RsTerminalGate0Result(
        RS_TERMINAL_GATE0_CALL*)(
            uint32_t requestedVersion,
            uint8_t* buffer,
            uint64_t capacity,
            uint64_t* outRequired);

inline constexpr char kRsTerminalGate0QueryLayoutJcsSymbol[] =
    "rs_terminal_gate0_query_layout_jcs";
inline constexpr uint32_t kRsTerminalGate0LayoutQueryVersion = 1;

namespace RedSalamander::TerminalEngine::Gate0::Layout
{
namespace Detail
{
#if defined(_M_X64)
inline constexpr std::string_view kArchitectureName = "x64";
#elif defined(_M_ARM64)
inline constexpr std::string_view kArchitectureName = "ARM64";
#else
#error The terminal Gate-0 layout proof supports only x64 and ARM64.
#endif

inline void AppendUnsigned(std::string& output, uint64_t value)
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

inline void AppendJcsString(
    std::string& output,
    std::string_view value)
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

struct EnumValue
{
    std::string_view name;
    uint64_t value;
};

template<typename Enum, size_t Count>
inline void AppendEnum(
    std::string& output,
    bool& first,
    std::string_view name,
    const std::array<EnumValue, Count>& values)
{
    if (!first)
    {
        output.push_back(',');
    }
    first = false;
    output.append("{\"align\":");
    AppendUnsigned(output, alignof(Enum));
    output.append(",\"name\":");
    AppendJcsString(output, name);
    output.append(",\"size\":");
    AppendUnsigned(output, sizeof(Enum));
    output.append(",\"values\":[");
    for (size_t index = 0; index < values.size(); ++index)
    {
        if (index != 0)
        {
            output.push_back(',');
        }
        output.append("{\"name\":");
        AppendJcsString(output, values[index].name);
        output.append(",\"value\":");
        AppendUnsigned(output, values[index].value);
        output.push_back('}');
    }
    output.append("]}");
}

inline void AppendField(
    std::string& output,
    bool& first,
    std::string_view name,
    size_t offset,
    size_t size,
    size_t alignment)
{
    if (!first)
    {
        output.push_back(',');
    }
    first = false;
    output.append("{\"align\":");
    AppendUnsigned(output, alignment);
    output.append(",\"name\":");
    AppendJcsString(output, name);
    output.append(",\"offset\":");
    AppendUnsigned(output, offset);
    output.append(",\"size\":");
    AppendUnsigned(output, size);
    output.push_back('}');
}

#define RS_GATE0_LAYOUT_APPEND_FIELD(output, first, type, field)         \
    AppendField(                                                         \
        output,                                                          \
        first,                                                           \
        #field,                                                          \
        offsetof(type, field),                                          \
        sizeof(((type*)nullptr)->field),                                 \
        alignof(decltype(((type*)nullptr)->field)))

template<typename Type, typename AppendFields>
inline void AppendStruct(
    std::string& output,
    bool& first,
    std::string_view name,
    AppendFields appendFields)
{
    if (!first)
    {
        output.push_back(',');
    }
    first = false;
    output.append("{\"align\":");
    AppendUnsigned(output, alignof(Type));
    output.append(",\"fields\":[");
    bool firstField = true;
    appendFields(firstField);
    output.append("],\"name\":");
    AppendJcsString(output, name);
    output.append(",\"size\":");
    AppendUnsigned(output, sizeof(Type));
    output.push_back('}');
}

template<typename Type>
inline void AppendFunctionPointer(
    std::string& output,
    bool& first,
    std::string_view name)
{
    if (!first)
    {
        output.push_back(',');
    }
    first = false;
    output.append("{\"align\":");
    AppendUnsigned(output, alignof(Type));
    output.append(",\"name\":");
    AppendJcsString(output, name);
    output.append(",\"size\":");
    AppendUnsigned(output, sizeof(Type));
    output.push_back('}');
}
} // namespace Detail

[[nodiscard]] inline std::string BuildLayoutJcs()
{
    using namespace Detail;

    std::string output;
    output.reserve(8192);
    output.append("{\"architecture\":");
    AppendJcsString(output, kArchitectureName);
    output.append(",\"callingConvention\":\"cdecl\",\"constants\":[");
    output.append(
        "{\"name\":\"RS_TERMINAL_GATE0_ADAPTER_ABI_V1\",\"value\":");
    AppendUnsigned(output, RS_TERMINAL_GATE0_ADAPTER_ABI_V1);
    output.append(
        "},{\"name\":\"RS_TERMINAL_GATE0_QUERY_SYMBOL_VERSION\","
        "\"value\":");
    AppendUnsigned(output, RS_TERMINAL_GATE0_QUERY_SYMBOL_VERSION);
    output.append(
        "},{\"name\":\"RS_TERMINAL_GATE0_LAYOUT_QUERY_VERSION\","
        "\"value\":");
    AppendUnsigned(output, kRsTerminalGate0LayoutQueryVersion);
    output.append("}],\"enums\":[");

    bool firstEnum = true;
    AppendEnum<RsTerminalGate0Result>(
        output,
        firstEnum,
        "RsTerminalGate0Result",
        std::array{
            EnumValue{"RS_TERMINAL_GATE0_SUCCESS",
                      RS_TERMINAL_GATE0_SUCCESS},
            EnumValue{"RS_TERMINAL_GATE0_INVALID_ARGUMENT",
                      RS_TERMINAL_GATE0_INVALID_ARGUMENT},
            EnumValue{"RS_TERMINAL_GATE0_UNSUPPORTED_ABI",
                      RS_TERMINAL_GATE0_UNSUPPORTED_ABI},
            EnumValue{"RS_TERMINAL_GATE0_OUT_OF_MEMORY",
                      RS_TERMINAL_GATE0_OUT_OF_MEMORY},
            EnumValue{"RS_TERMINAL_GATE0_OUT_OF_SPACE",
                      RS_TERMINAL_GATE0_OUT_OF_SPACE},
            EnumValue{"RS_TERMINAL_GATE0_LIMIT_REJECTED",
                      RS_TERMINAL_GATE0_LIMIT_REJECTED},
            EnumValue{"RS_TERMINAL_GATE0_INTERNAL_ERROR",
                      RS_TERMINAL_GATE0_INTERNAL_ERROR},
        });
    AppendEnum<RsTerminalGate0Architecture>(
        output,
        firstEnum,
        "RsTerminalGate0Architecture",
        std::array{
            EnumValue{"RS_TERMINAL_GATE0_ARCH_UNKNOWN",
                      RS_TERMINAL_GATE0_ARCH_UNKNOWN},
            EnumValue{"RS_TERMINAL_GATE0_ARCH_X64",
                      RS_TERMINAL_GATE0_ARCH_X64},
            EnumValue{"RS_TERMINAL_GATE0_ARCH_ARM64",
                      RS_TERMINAL_GATE0_ARCH_ARM64},
        });
    AppendEnum<RsTerminalGate0CallingConvention>(
        output,
        firstEnum,
        "RsTerminalGate0CallingConvention",
        std::array{
            EnumValue{"RS_TERMINAL_GATE0_CALL_UNKNOWN",
                      RS_TERMINAL_GATE0_CALL_UNKNOWN},
            EnumValue{"RS_TERMINAL_GATE0_CALL_CDECL",
                      RS_TERMINAL_GATE0_CALL_CDECL},
        });
    AppendEnum<RsTerminalGate0ResourceKind>(
        output,
        firstEnum,
        "RsTerminalGate0ResourceKind",
        std::array{
            EnumValue{"RS_TERMINAL_GATE0_RESOURCE_TERMINAL_TEXT_CPU",
                      RS_TERMINAL_GATE0_RESOURCE_TERMINAL_TEXT_CPU},
            EnumValue{"RS_TERMINAL_GATE0_RESOURCE_IMAGE_CPU",
                      RS_TERMINAL_GATE0_RESOURCE_IMAGE_CPU},
            EnumValue{"RS_TERMINAL_GATE0_RESOURCE_IMAGE_GPU",
                      RS_TERMINAL_GATE0_RESOURCE_IMAGE_GPU},
            EnumValue{"RS_TERMINAL_GATE0_RESOURCE_PENDING_EFFECT",
                      RS_TERMINAL_GATE0_RESOURCE_PENDING_EFFECT},
            EnumValue{"RS_TERMINAL_GATE0_RESOURCE_COUNT",
                      RS_TERMINAL_GATE0_RESOURCE_COUNT},
        });
    AppendEnum<RsTerminalGate0Capability>(
        output,
        firstEnum,
        "RsTerminalGate0Capability",
        std::array{
            EnumValue{"RS_TERMINAL_GATE0_CAP_CELLS_AND_STYLES",
                      RS_TERMINAL_GATE0_CAP_CELLS_AND_STYLES},
            EnumValue{"RS_TERMINAL_GATE0_CAP_SNAPSHOT_COPY",
                      RS_TERMINAL_GATE0_CAP_SNAPSHOT_COPY},
            EnumValue{"RS_TERMINAL_GATE0_CAP_INPUT_ENCODING",
                      RS_TERMINAL_GATE0_CAP_INPUT_ENCODING},
            EnumValue{"RS_TERMINAL_GATE0_CAP_EFFECT_ORDERING",
                      RS_TERMINAL_GATE0_CAP_EFFECT_ORDERING},
            EnumValue{"RS_TERMINAL_GATE0_CAP_KITTY_STATIC",
                      RS_TERMINAL_GATE0_CAP_KITTY_STATIC},
            EnumValue{"RS_TERMINAL_GATE0_CAP_RESOURCE_ACCOUNTING",
                      RS_TERMINAL_GATE0_CAP_RESOURCE_ACCOUNTING},
            EnumValue{"RS_TERMINAL_GATE0_CAP_QUIET_DESTROY",
                      RS_TERMINAL_GATE0_CAP_QUIET_DESTROY},
        });
    AppendEnum<RsTerminalGate0ActionFlag>(
        output,
        firstEnum,
        "RsTerminalGate0ActionFlag",
        std::array{
            EnumValue{"RS_TERMINAL_GATE0_ACTION_OBSERVED",
                      RS_TERMINAL_GATE0_ACTION_OBSERVED},
            EnumValue{"RS_TERMINAL_GATE0_ACTION_CHECKED_ARITHMETIC",
                      RS_TERMINAL_GATE0_ACTION_CHECKED_ARITHMETIC},
            EnumValue{"RS_TERMINAL_GATE0_ACTION_BOUNDARY_ACCEPTED",
                      RS_TERMINAL_GATE0_ACTION_BOUNDARY_ACCEPTED},
            EnumValue{"RS_TERMINAL_GATE0_ACTION_PLUS_ONE_REJECTED",
                      RS_TERMINAL_GATE0_ACTION_PLUS_ONE_REJECTED},
            EnumValue{
                "RS_TERMINAL_GATE0_ACTION_REJECTION_BEFORE_ALLOCATION",
                RS_TERMINAL_GATE0_ACTION_REJECTION_BEFORE_ALLOCATION},
            EnumValue{"RS_TERMINAL_GATE0_ACTION_NO_PARTIAL_EFFECT",
                      RS_TERMINAL_GATE0_ACTION_NO_PARTIAL_EFFECT},
            EnumValue{"RS_TERMINAL_GATE0_ACTION_ORDER_PRESERVED",
                      RS_TERMINAL_GATE0_ACTION_ORDER_PRESERVED},
            EnumValue{"RS_TERMINAL_GATE0_ACTION_QUIET_POINT_REACHED",
                      RS_TERMINAL_GATE0_ACTION_QUIET_POINT_REACHED},
        });

    output.append("],\"functionPointers\":[");
    bool firstFunction = true;
    AppendFunctionPointer<RsTerminalGate0ReserveResourceFn>(
        output,
        firstFunction,
        "RsTerminalGate0ReserveResourceFn");
    AppendFunctionPointer<RsTerminalGate0ReleaseResourceFn>(
        output,
        firstFunction,
        "RsTerminalGate0ReleaseResourceFn");
    AppendFunctionPointer<RsTerminalGate0SessionCreateFn>(
        output,
        firstFunction,
        "RsTerminalGate0SessionCreateFn");
    AppendFunctionPointer<RsTerminalGate0SessionWriteFn>(
        output,
        firstFunction,
        "RsTerminalGate0SessionWriteFn");
    AppendFunctionPointer<RsTerminalGate0SessionActionFn>(
        output,
        firstFunction,
        "RsTerminalGate0SessionActionFn");
    AppendFunctionPointer<RsTerminalGate0SessionStateJcsFn>(
        output,
        firstFunction,
        "RsTerminalGate0SessionStateJcsFn");
    AppendFunctionPointer<RsTerminalGate0SessionDestroyFn>(
        output,
        firstFunction,
        "RsTerminalGate0SessionDestroyFn");
    AppendFunctionPointer<RsTerminalGate0QueryAdapterFn>(
        output,
        firstFunction,
        "RsTerminalGate0QueryAdapterFn");
    AppendFunctionPointer<RsTerminalGate0QueryLayoutJcsFn>(
        output,
        firstFunction,
        "RsTerminalGate0QueryLayoutJcsFn");

    output.append("],\"pointerSize\":");
    AppendUnsigned(output, sizeof(void*));
    output.append(",\"structs\":[");
    bool firstStruct = true;
    AppendStruct<RsTerminalGate0String>(
        output,
        firstStruct,
        "RsTerminalGate0String",
        [&](bool& first) {
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0String,
                ptr);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0String,
                len);
        });
    AppendStruct<RsTerminalGate0IdentityV1>(
        output,
        firstStruct,
        "RsTerminalGate0IdentityV1",
        [&](bool& first) {
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0IdentityV1,
                size);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0IdentityV1,
                adapter_abi);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0IdentityV1,
                candidate_id);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0IdentityV1,
                upstream_pin);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0IdentityV1,
                source_sha256);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0IdentityV1,
                adapter_header_sha256);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0IdentityV1,
                build_identity);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0IdentityV1,
                architecture);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0IdentityV1,
                calling_convention);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0IdentityV1,
                capabilities);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0IdentityV1,
                pointer_size);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0IdentityV1,
                reserved);
        });
    AppendStruct<RsTerminalGate0ResourceLimitsV1>(
        output,
        firstStruct,
        "RsTerminalGate0ResourceLimitsV1",
        [&](bool& first) {
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceLimitsV1,
                size);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceLimitsV1,
                terminal_text_session_bytes);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceLimitsV1,
                terminal_text_process_bytes);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceLimitsV1,
                image_cpu_session_bytes);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceLimitsV1,
                image_cpu_process_bytes);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceLimitsV1,
                image_gpu_session_bytes);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceLimitsV1,
                image_gpu_process_bytes);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceLimitsV1,
                pending_effect_session_bytes);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceLimitsV1,
                pending_effect_process_bytes);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceLimitsV1,
                pending_effect_session_descriptors);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceLimitsV1,
                pending_effect_process_descriptors);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceLimitsV1,
                maximum_sessions);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceLimitsV1,
                kitty_max_pixels);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceLimitsV1,
                kitty_max_dimension);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceLimitsV1,
                reserved);
        });
    AppendStruct<RsTerminalGate0ResourceCallbacksV1>(
        output,
        firstStruct,
        "RsTerminalGate0ResourceCallbacksV1",
        [&](bool& first) {
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceCallbacksV1,
                size);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceCallbacksV1,
                userdata);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceCallbacksV1,
                reserve);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ResourceCallbacksV1,
                release);
        });
    AppendStruct<RsTerminalGate0SessionConfigV1>(
        output,
        firstStruct,
        "RsTerminalGate0SessionConfigV1",
        [&](bool& first) {
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0SessionConfigV1,
                size);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0SessionConfigV1,
                columns);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0SessionConfigV1,
                rows);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0SessionConfigV1,
                cell_width_px);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0SessionConfigV1,
                cell_height_px);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0SessionConfigV1,
                fixture_id);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0SessionConfigV1,
                schedule_id);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0SessionConfigV1,
                limits);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0SessionConfigV1,
                resource_callbacks);
        });
    AppendStruct<RsTerminalGate0ActionV1>(
        output,
        firstStruct,
        "RsTerminalGate0ActionV1",
        [&](bool& first) {
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ActionV1,
                size);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ActionV1,
                action_id);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ActionV1,
                arguments_jcs);
        });
    AppendStruct<RsTerminalGate0ActionObservationV1>(
        output,
        firstStruct,
        "RsTerminalGate0ActionObservationV1",
        [&](bool& first) {
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ActionObservationV1,
                size);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ActionObservationV1,
                flags);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ActionObservationV1,
                reservation_calls_before_rejection);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0ActionObservationV1,
                allocation_attempts_before_rejection);
        });
    AppendStruct<RsTerminalGate0AdapterV1>(
        output,
        firstStruct,
        "RsTerminalGate0AdapterV1",
        [&](bool& first) {
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0AdapterV1,
                size);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0AdapterV1,
                identity);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0AdapterV1,
                session_create);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0AdapterV1,
                session_write);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0AdapterV1,
                session_action);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0AdapterV1,
                session_state_jcs);
            RS_GATE0_LAYOUT_APPEND_FIELD(
                output,
                first,
                RsTerminalGate0AdapterV1,
                session_destroy);
        });
    output.append("],\"version\":1}");
    return output;
}

[[nodiscard]] inline RsTerminalGate0Result QueryLayoutJcs(
    uint32_t requestedVersion,
    uint8_t* buffer,
    uint64_t capacity,
    uint64_t* outRequired) noexcept
{
    if (requestedVersion != kRsTerminalGate0LayoutQueryVersion)
    {
        return RS_TERMINAL_GATE0_UNSUPPORTED_ABI;
    }
    if (outRequired == nullptr ||
        (buffer == nullptr && capacity != 0))
    {
        return RS_TERMINAL_GATE0_INVALID_ARGUMENT;
    }

    try
    {
        const std::string layout = BuildLayoutJcs();
        *outRequired = static_cast<uint64_t>(layout.size());
        if (buffer == nullptr)
        {
            return RS_TERMINAL_GATE0_SUCCESS;
        }
        if (capacity < static_cast<uint64_t>(layout.size()))
        {
            return RS_TERMINAL_GATE0_OUT_OF_SPACE;
        }
        std::memcpy(buffer, layout.data(), layout.size());
        return RS_TERMINAL_GATE0_SUCCESS;
    }
    catch (const std::bad_alloc&)
    {
        // Allocation failure at this required C ABI proof seam is fatal.
        std::terminate();
    }
    catch (const std::length_error&)
    {
        // Fixed schema growth cannot legitimately overflow std::string.
        return RS_TERMINAL_GATE0_INTERNAL_ERROR;
    }
}
} // namespace RedSalamander::TerminalEngine::Gate0::Layout

#undef RS_GATE0_LAYOUT_APPEND_FIELD

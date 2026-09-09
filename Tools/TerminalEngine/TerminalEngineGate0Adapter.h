#pragma once

// Candidate-neutral Gate-0 terminal-engine adapter ABI.
//
// This header is consumed by the evidence harness and by candidate-specific
// spike adapters only. It is not a product plugin ABI. All pointers borrowed
// from the adapter remain valid until the adapter DLL is unloaded. A session
// owns no host memory after session_destroy returns.

#include <stdbool.h>
#include <stddef.h>
#include <stdint.h>

#if defined(_WIN32)
#define RS_TERMINAL_GATE0_CALL __cdecl
#if defined(RS_TERMINAL_GATE0_ADAPTER_EXPORTS)
#define RS_TERMINAL_GATE0_API __declspec(dllexport)
#else
#define RS_TERMINAL_GATE0_API __declspec(dllimport)
#endif
#else
#define RS_TERMINAL_GATE0_CALL
#define RS_TERMINAL_GATE0_API
#endif

#ifdef __cplusplus
extern "C"
{
#endif

    enum
    {
        RS_TERMINAL_GATE0_ADAPTER_ABI_V1       = 1,
        RS_TERMINAL_GATE0_QUERY_SYMBOL_VERSION = 1,
    };

    typedef enum RsTerminalGate0Result
    {
        RS_TERMINAL_GATE0_SUCCESS          = 0,
        RS_TERMINAL_GATE0_INVALID_ARGUMENT = 1,
        RS_TERMINAL_GATE0_UNSUPPORTED_ABI  = 2,
        RS_TERMINAL_GATE0_OUT_OF_MEMORY    = 3,
        RS_TERMINAL_GATE0_OUT_OF_SPACE     = 4,
        RS_TERMINAL_GATE0_LIMIT_REJECTED   = 5,
        RS_TERMINAL_GATE0_INTERNAL_ERROR   = 6,
    } RsTerminalGate0Result;

    typedef enum RsTerminalGate0Architecture
    {
        RS_TERMINAL_GATE0_ARCH_UNKNOWN = 0,
        RS_TERMINAL_GATE0_ARCH_X64     = 1,
        RS_TERMINAL_GATE0_ARCH_ARM64   = 2,
    } RsTerminalGate0Architecture;

    typedef enum RsTerminalGate0CallingConvention
    {
        RS_TERMINAL_GATE0_CALL_UNKNOWN = 0,
        RS_TERMINAL_GATE0_CALL_CDECL   = 1,
    } RsTerminalGate0CallingConvention;

    typedef enum RsTerminalGate0ResourceKind
    {
        RS_TERMINAL_GATE0_RESOURCE_TERMINAL_TEXT_CPU = 0,
        RS_TERMINAL_GATE0_RESOURCE_IMAGE_CPU         = 1,
        RS_TERMINAL_GATE0_RESOURCE_IMAGE_GPU         = 2,
        RS_TERMINAL_GATE0_RESOURCE_PENDING_EFFECT    = 3,
        RS_TERMINAL_GATE0_RESOURCE_COUNT             = 4,
    } RsTerminalGate0ResourceKind;

    typedef enum RsTerminalGate0Capability
    {
        RS_TERMINAL_GATE0_CAP_CELLS_AND_STYLES    = UINT64_C(1) << 0,
        RS_TERMINAL_GATE0_CAP_SNAPSHOT_COPY       = UINT64_C(1) << 1,
        RS_TERMINAL_GATE0_CAP_INPUT_ENCODING      = UINT64_C(1) << 2,
        RS_TERMINAL_GATE0_CAP_EFFECT_ORDERING     = UINT64_C(1) << 3,
        RS_TERMINAL_GATE0_CAP_KITTY_STATIC        = UINT64_C(1) << 4,
        RS_TERMINAL_GATE0_CAP_RESOURCE_ACCOUNTING = UINT64_C(1) << 5,
        RS_TERMINAL_GATE0_CAP_QUIET_DESTROY       = UINT64_C(1) << 6,
    } RsTerminalGate0Capability;

    typedef enum RsTerminalGate0ActionFlag
    {
        RS_TERMINAL_GATE0_ACTION_OBSERVED                    = UINT64_C(1) << 0,
        RS_TERMINAL_GATE0_ACTION_CHECKED_ARITHMETIC          = UINT64_C(1) << 1,
        RS_TERMINAL_GATE0_ACTION_BOUNDARY_ACCEPTED           = UINT64_C(1) << 2,
        RS_TERMINAL_GATE0_ACTION_PLUS_ONE_REJECTED           = UINT64_C(1) << 3,
        RS_TERMINAL_GATE0_ACTION_REJECTION_BEFORE_ALLOCATION = UINT64_C(1) << 4,
        RS_TERMINAL_GATE0_ACTION_NO_PARTIAL_EFFECT           = UINT64_C(1) << 5,
        RS_TERMINAL_GATE0_ACTION_ORDER_PRESERVED             = UINT64_C(1) << 6,
        RS_TERMINAL_GATE0_ACTION_QUIET_POINT_REACHED         = UINT64_C(1) << 7,
    } RsTerminalGate0ActionFlag;

    typedef struct RsTerminalGate0String
    {
        const uint8_t* ptr;
        size_t len;
    } RsTerminalGate0String;

    typedef struct RsTerminalGate0IdentityV1
    {
        size_t size;
        uint32_t adapter_abi;
        RsTerminalGate0String candidate_id;
        RsTerminalGate0String upstream_pin;
        RsTerminalGate0String source_sha256;
        RsTerminalGate0String adapter_header_sha256;
        RsTerminalGate0String build_identity;
        RsTerminalGate0Architecture architecture;
        RsTerminalGate0CallingConvention calling_convention;
        uint64_t capabilities;
        uint32_t pointer_size;
        uint32_t reserved;
    } RsTerminalGate0IdentityV1;

    typedef struct RsTerminalGate0ResourceLimitsV1
    {
        size_t size;
        uint64_t terminal_text_session_bytes;
        uint64_t terminal_text_process_bytes;
        uint64_t image_cpu_session_bytes;
        uint64_t image_cpu_process_bytes;
        uint64_t image_gpu_session_bytes;
        uint64_t image_gpu_process_bytes;
        uint64_t pending_effect_session_bytes;
        uint64_t pending_effect_process_bytes;
        uint32_t pending_effect_session_descriptors;
        uint32_t pending_effect_process_descriptors;
        uint32_t maximum_sessions;
        uint32_t kitty_max_pixels;
        uint32_t kitty_max_dimension;
        uint32_t reserved;
    } RsTerminalGate0ResourceLimitsV1;

    typedef bool(RS_TERMINAL_GATE0_CALL* RsTerminalGate0ReserveResourceFn)(void* userdata,
                                                                           RsTerminalGate0ResourceKind kind,
                                                                           uint64_t bytes,
                                                                           uint32_t descriptors);

    typedef void(RS_TERMINAL_GATE0_CALL* RsTerminalGate0ReleaseResourceFn)(void* userdata,
                                                                           RsTerminalGate0ResourceKind kind,
                                                                           uint64_t bytes,
                                                                           uint32_t descriptors);

    typedef struct RsTerminalGate0ResourceCallbacksV1
    {
        size_t size;
        void* userdata;
        RsTerminalGate0ReserveResourceFn reserve;
        RsTerminalGate0ReleaseResourceFn release;
    } RsTerminalGate0ResourceCallbacksV1;

    typedef struct RsTerminalGate0SessionConfigV1
    {
        size_t size;
        uint16_t columns;
        uint16_t rows;
        uint16_t cell_width_px;
        uint16_t cell_height_px;
        RsTerminalGate0String fixture_id;
        RsTerminalGate0String schedule_id;
        const RsTerminalGate0ResourceLimitsV1* limits;
        const RsTerminalGate0ResourceCallbacksV1* resource_callbacks;
    } RsTerminalGate0SessionConfigV1;

    typedef struct RsTerminalGate0ActionV1
    {
        size_t size;
        RsTerminalGate0String action_id;
        RsTerminalGate0String arguments_jcs;
    } RsTerminalGate0ActionV1;

    typedef struct RsTerminalGate0ActionObservationV1
    {
        size_t size;
        uint64_t flags;
        uint64_t reservation_calls_before_rejection;
        uint64_t allocation_attempts_before_rejection;
    } RsTerminalGate0ActionObservationV1;

    typedef void* RsTerminalGate0Session;

    typedef RsTerminalGate0Result(RS_TERMINAL_GATE0_CALL* RsTerminalGate0SessionCreateFn)(const RsTerminalGate0SessionConfigV1* config,
                                                                                          RsTerminalGate0Session* out_session);

    typedef RsTerminalGate0Result(RS_TERMINAL_GATE0_CALL* RsTerminalGate0SessionWriteFn)(RsTerminalGate0Session session, const uint8_t* data, size_t length);

    typedef RsTerminalGate0Result(RS_TERMINAL_GATE0_CALL* RsTerminalGate0SessionActionFn)(RsTerminalGate0Session session,
                                                                                          const RsTerminalGate0ActionV1* action,
                                                                                          RsTerminalGate0ActionObservationV1* out_observation);

    typedef RsTerminalGate0Result(RS_TERMINAL_GATE0_CALL* RsTerminalGate0SessionStateJcsFn)(RsTerminalGate0Session session,
                                                                                            uint8_t* buffer,
                                                                                            size_t capacity,
                                                                                            size_t* out_required);

    typedef RsTerminalGate0Result(RS_TERMINAL_GATE0_CALL* RsTerminalGate0SessionDestroyFn)(RsTerminalGate0Session session);

    typedef struct RsTerminalGate0AdapterV1
    {
        size_t size;
        RsTerminalGate0IdentityV1 identity;
        RsTerminalGate0SessionCreateFn session_create;
        RsTerminalGate0SessionWriteFn session_write;
        RsTerminalGate0SessionActionFn session_action;
        RsTerminalGate0SessionStateJcsFn session_state_jcs;
        RsTerminalGate0SessionDestroyFn session_destroy;
    } RsTerminalGate0AdapterV1;

    typedef RsTerminalGate0Result(RS_TERMINAL_GATE0_CALL* RsTerminalGate0QueryAdapterFn)(uint32_t query_symbol_version,
                                                                                         uint32_t requested_adapter_abi,
                                                                                         RsTerminalGate0AdapterV1* out_adapter);

    RS_TERMINAL_GATE0_API RsTerminalGate0Result RS_TERMINAL_GATE0_CALL rs_terminal_gate0_query_adapter(uint32_t query_symbol_version,
                                                                                                       uint32_t requested_adapter_abi,
                                                                                                       RsTerminalGate0AdapterV1* out_adapter);

#ifdef __cplusplus
} // extern "C"
#endif

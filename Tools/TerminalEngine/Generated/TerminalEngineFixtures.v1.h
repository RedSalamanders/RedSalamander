#pragma once

#include <array>
#include <cstddef>
#include <cstdint>
#include <span>
#include <string_view>

namespace RedSalamander::TerminalEngine::Fixtures::V1 {
enum class InputOp : std::uint8_t { LiteralHex, RepeatByte };
struct InputStep final { InputOp op; std::string_view hex; std::uint8_t byteValue; std::uint64_t repeatCount; };
enum class ScheduleAlgorithm : std::uint8_t { AllAtOnce, Bytewise, Fibonacci, TerminatorCuts };
struct ScheduleSpec final { std::string_view scheduleId; ScheduleAlgorithm algorithm; std::span<const std::uint64_t> cutOffsets; std::string_view scheduleSha256; };
struct Check final { std::string_view checkId; std::string_view expectedValue; std::string_view expectedSha256; };
struct Action final { std::string_view actionId; std::string_view argumentsJcs; std::string_view actionSha256; };
struct Fixture final { std::string_view fixtureId; std::span<const InputStep> inputPlan; std::uint64_t inputSizeBytes; std::string_view inputSha256; std::span<const ScheduleSpec> schedules; std::span<const Action> actions; std::span<const Check> checks; std::string_view contractSha256; };
struct PerformancePayloadSpec final { std::span<const std::string_view> fixtureIds; std::uint64_t sizeBytes; std::string_view sha256; };

inline constexpr std::array<InputStep, 1> k_alternate_screen_roundtrip_input{
    InputStep{InputOp::LiteralHex, "7072696d6172791b5b3f3130343968616c7465726e6174651b5b3f313034396c", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 2> k_alternate_screen_roundtrip_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "e994c1fff15b0e175687114a8263b98e6a7fa7b5265f53e120045e99472048b1"},
    ScheduleSpec{"bytewise", ScheduleAlgorithm::Bytewise, std::span<const std::uint64_t>{}, "e4f21a3f295b96a38c8a7376c7760cd400889054e1713696882a03d499870836"},
};
inline constexpr std::array<Check, 2> k_alternate_screen_roundtrip_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"alternateDiscarded\":true,\"primarySurvived\":true,\"primaryText\":\"primary\"}", "340fc0ff169251e3d225c2e1f3579cb834f7a75ec787a4eee1cf3bd5efe0a515"},
};
inline constexpr std::array<Action, 1> k_alternate_screen_roundtrip_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
};
inline constexpr std::array<InputStep, 1> k_bracketed_paste_modes_input{
    InputStep{InputOp::LiteralHex, "1b5b3f3230303468", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 1> k_bracketed_paste_modes_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "4f8241d07df07bf53ffada796118dd8a91756e566ab7d29faf909c247954dac4"},
};
inline constexpr std::array<Check, 2> k_bracketed_paste_modes_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"disabled\":\"payload\",\"enabled\":\"\\\\u001b[200~payload\\\\u001b[201~\",\"ordered\":true}", "b6e9ddad790247d00e2c965abefa78d4e47c7e3936d3c0b7a9d9cbd831e1e592"},
};
inline constexpr std::array<Action, 3> k_bracketed_paste_modes_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
    Action{"encode-paste-disabled", "{\"payloadHex\":\"7061796c6f6164\"}", "c8216be8715a6e94c56d1027c2c96c0ed7adc431547c7f7bd0615f53fc2c4f90"},
    Action{"encode-paste-enabled", "{\"payloadHex\":\"7061796c6f6164\"}", "5ad390e75e703df689333d5fa36717f4266ab892ae08fcfe2ccff131f1ffde92"},
};
inline constexpr std::array<InputStep, 1> k_color_style_matrix_input{
    InputStep{InputOp::LiteralHex, "1b5b313b323b333b343a333b353b373b393b33313b34383b353b3230306d411b5b33383b323b313b323b333b34383b323b343b353b366d421b5b306d43", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 1> k_color_style_matrix_schedules{
    ScheduleSpec{"fibonacci", ScheduleAlgorithm::Fibonacci, std::span<const std::uint64_t>{}, "00a608a31eb4321dc6bc3e1b67788aa3a02dfe2971f591beeb02bdd6f651d1b1"},
};
inline constexpr std::array<Check, 2> k_color_style_matrix_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"indexed\":true,\"styles\":\"bold,faint,italic,underline,blink,inverse,strike\",\"truecolor\":\"1,2,3/4,5,6\"}", "34850d447d1373c4ef9e1b14821df9a6d486dea1c1b6c67d896236426458f1a7"},
};
inline constexpr std::array<Action, 1> k_color_style_matrix_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
};
inline constexpr std::array<InputStep, 1> k_cursor_and_palette_input{
    InputStep{InputOp::LiteralHex, "1b5b353b37481b5b3f32356c1b5d343b313b7267623a31312f32322f33331b5c1b5d31303b7267623a34342f35352f36361b5c", 0, 0ULL},
};
inline constexpr std::array<std::uint64_t, 13> k_cursor_and_palette_terminator_split_cuts{1ULL, 5ULL, 6ULL, 7ULL, 11ULL, 12ULL, 13ULL, 30ULL, 31ULL, 32ULL, 33ULL, 49ULL, 50ULL};
inline constexpr std::array<ScheduleSpec, 1> k_cursor_and_palette_schedules{
    ScheduleSpec{"terminator-split", ScheduleAlgorithm::TerminatorCuts, k_cursor_and_palette_terminator_split_cuts, "6c6543d0c1fdb62cb8e3516d3766310b26c9c76ce89e148c775e78de76646ff2"},
};
inline constexpr std::array<Check, 2> k_cursor_and_palette_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"cursorVisible\":false,\"defaultForeground\":\"44,55,66\",\"palette1\":\"11,22,33\"}", "378baaf7d249e755502d82d4eb4587a78f250fce753d2a380dec35141cfdb211"},
};
inline constexpr std::array<Action, 1> k_cursor_and_palette_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
};
inline constexpr std::array<InputStep, 1> k_focus_and_mouse_matrix_input{
    InputStep{InputOp::LiteralHex, "1b5b3f31303030681b5b3f31303032681b5b3f31303033681b5b3f31303035681b5b3f31303036681b5b3f31303135681b5b3f31303136681b5b3f3130303468", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 1> k_focus_and_mouse_matrix_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "9fdf5b418f52098cd3c37967a2b09c7b3b3394dcc58d3c0277904a346c6a8295"},
};
inline constexpr std::array<Check, 2> k_focus_and_mouse_matrix_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"focus\":true,\"modes\":\"1000,1002,1003,1005,1006,1015,1016\",\"ordered\":true}", "66e4c5c16e3d202dbf6eb123b818fcb91d2c6e230fab60939a59d1a9d7be1849"},
};
inline constexpr std::array<Action, 3> k_focus_and_mouse_matrix_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
    Action{"encode-focus-gained-lost", "{\"order\":\"gained,lost\"}", "9c25dab79d33de20cba37b61ae4737dd2e6a22b1a5d34edce0c2e620b273f658"},
    Action{"encode-mouse-matrix", "{\"modes\":\"1000,1002,1003,1005,1006,1015,1016\"}", "7da4d789c4efe10a52d1bf59bcae0ee1b9d96ad27a3081119b94b970da0c3331"},
};
inline constexpr std::array<InputStep, 11> k_hyperlink_boundaries_input{
    InputStep{InputOp::LiteralHex, "1b5d383b", 0, 0ULL},
    InputStep{InputOp::RepeatByte, "", 112, 1024ULL},
    InputStep{InputOp::LiteralHex, "3b75", 0, 0ULL},
    InputStep{InputOp::LiteralHex, "1b5c701b5d383b", 0, 0ULL},
    InputStep{InputOp::RepeatByte, "", 113, 1025ULL},
    InputStep{InputOp::LiteralHex, "3b75", 0, 0ULL},
    InputStep{InputOp::LiteralHex, "1b5c711b5d383b3b", 0, 0ULL},
    InputStep{InputOp::RepeatByte, "", 117, 8192ULL},
    InputStep{InputOp::LiteralHex, "1b5c751b5d383b3b", 0, 0ULL},
    InputStep{InputOp::RepeatByte, "", 118, 8193ULL},
    InputStep{InputOp::LiteralHex, "1b5c76", 0, 0ULL},
};
inline constexpr std::array<std::uint64_t, 15> k_hyperlink_boundaries_terminator_split_cuts{1ULL, 1030ULL, 1031ULL, 1033ULL, 1034ULL, 2064ULL, 2065ULL, 2067ULL, 2068ULL, 10264ULL, 10265ULL, 10267ULL, 10268ULL, 18465ULL, 18466ULL};
inline constexpr std::array<ScheduleSpec, 3> k_hyperlink_boundaries_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "059cb5d4563958039dfb1518ecbb63ea17841cd99ad2761b8d064d2d9a50f40f"},
    ScheduleSpec{"bytewise", ScheduleAlgorithm::Bytewise, std::span<const std::uint64_t>{}, "159b5c090b1fbb4f0187f81682422e17fe290d03cf512258e4c590e37dd6aff9"},
    ScheduleSpec{"terminator-split", ScheduleAlgorithm::TerminatorCuts, k_hyperlink_boundaries_terminator_split_cuts, "159c21ade7c7164d0f24cccfe413dbbd9673ca6280aaf373a2e8e6e092284ea9"},
};
inline constexpr std::array<Check, 2> k_hyperlink_boundaries_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":true,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":true}", "b87b21bd5ea0f04d2da3977bff3aa4f80086ecda8d182daa2b215e12e2f320f4"},
    Check{"state", "{\"parameterBoundary\":\"1024\",\"plusOneDiscarded\":true,\"recordBoundary\":\"16384\",\"uriBoundary\":\"8192\"}", "146f81e59c00d715bc1b644739533bb078cf63f510a892aecb7a076d26175b88"},
};
inline constexpr std::array<Action, 2> k_hyperlink_boundaries_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
    Action{"probe-osc8-limits", "{\"parameter\":\"1024\",\"record\":\"16384\",\"uri\":\"8192\"}", "a252614dd814cbcacba65e95107b19f9883b3dd1f4ba96a34b178df0688cbf0e"},
};
inline constexpr std::array<InputStep, 1> k_kitty_delete_generation_input{
    InputStep{InputOp::LiteralHex, "1b5f47663d33322c733d312c763d312c693d372c613d543b2f7741412f773d3d1b5c1b5f47613d702c693d372c703d31312c713d323b1b5c1b5f47613d642c643d692c693d373b1b5c", 0, 0ULL},
};
inline constexpr std::array<std::uint64_t, 11> k_kitty_delete_generation_terminator_split_cuts{1ULL, 32ULL, 33ULL, 34ULL, 35ULL, 54ULL, 55ULL, 56ULL, 57ULL, 71ULL, 72ULL};
inline constexpr std::array<ScheduleSpec, 2> k_kitty_delete_generation_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "8fc9a3529d935b88e504c6b70a94e4e2fab552e3677f274ae12916bfa3856f3f"},
    ScheduleSpec{"terminator-split", ScheduleAlgorithm::TerminatorCuts, k_kitty_delete_generation_terminator_split_cuts, "b0a2a9abdef01beabbc5d10011c56a3a2aa54637c54ba74256522da6c2d67d75"},
};
inline constexpr std::array<Check, 2> k_kitty_delete_generation_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"imageGenerationChanged\":true,\"placementGenerationChanged\":true,\"placementOnlyMutationDistinct\":true}", "f38190693dee59720adef915a286b493d465798516a7cd740fd05e54a9350ba2"},
};
inline constexpr std::array<Action, 1> k_kitty_delete_generation_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
};
inline constexpr std::array<InputStep, 1> k_kitty_dimension_overflow_input{
    InputStep{InputOp::LiteralHex, "1b5f47663d33322c733d343239343936373239352c763d343239343936373239352c693d382c613d543b41413d3d1b5c", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 1> k_kitty_dimension_overflow_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "41cc6a31fe1504b63b8e6f9d3687edcd9f9ae0c27d9d8d4737bb82b8f2291927"},
};
inline constexpr std::array<Check, 2> k_kitty_dimension_overflow_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":true,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":true}", "b87b21bd5ea0f04d2da3977bff3aa4f80086ecda8d182daa2b215e12e2f320f4"},
    Check{"state", "{\"allocationAttempted\":false,\"checkedArithmetic\":true,\"rejected\":true}", "5e9da6f8fa9a226bfa5c8587ea1465365ec1e63a2661110501904abf3acbb8d9"},
};
inline constexpr std::array<Action, 1> k_kitty_dimension_overflow_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
};
inline constexpr std::array<InputStep, 1> k_kitty_direct_png_input{
    InputStep{InputOp::LiteralHex, "1b5f47663d3130302c693d31332c613d543b6956424f5277304b47676f414141414e5355684555674141414145414141414243415941414141664663534a4141414144556c4551565234326d50347a384477487741464141482f56736376445141414141424a52553545726b4a6767673d3d1b5c", 0, 0ULL},
};
inline constexpr std::array<std::uint64_t, 3> k_kitty_direct_png_terminator_split_cuts{1ULL, 114ULL, 115ULL};
inline constexpr std::array<ScheduleSpec, 2> k_kitty_direct_png_schedules{
    ScheduleSpec{"fibonacci", ScheduleAlgorithm::Fibonacci, std::span<const std::uint64_t>{}, "03d20b4e9e94efba584c33bc54e3a7cdcc383a1917309039966df54b0ab0276d"},
    ScheduleSpec{"terminator-split", ScheduleAlgorithm::TerminatorCuts, k_kitty_direct_png_terminator_split_cuts, "ff5bf25332e2c5ad5d879c7ec942dd2f4e3c234302484297090c5e90b283d324"},
};
inline constexpr std::array<Check, 2> k_kitty_direct_png_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"format\":\"png\",\"height\":\"1\",\"imageId\":\"13\",\"immutableAfterDelete\":true,\"pixel\":\"ff0000ff\",\"snapshotOwned\":true,\"width\":\"1\"}", "64a4f66c4337d318dc879bda36c59a7f5db30d0f103d8f370444eae9407171d9"},
};
inline constexpr std::array<Action, 2> k_kitty_direct_png_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
    Action{"snapshot-static-kitty", "{\"expectedDecodedRgbaHex\":\"ff0000ff\",\"imageId\":\"13\",\"mutation\":\"delete-image-and-placement\",\"requireOwnedCopyStable\":true}", "4416bac3f552fc5782791a90dc20068b6a12df70900305a8dc47a14c7932fc88"},
};
inline constexpr std::array<InputStep, 1> k_kitty_direct_rgb_input{
    InputStep{InputOp::LiteralHex, "1b5f47663d32342c733d312c763d312c693d312c613d543b2f7741411b5c", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 2> k_kitty_direct_rgb_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "1635c31bcfe16560870201e84b80bcb77f0f37de7370eaf09a2f862cb9452a24"},
    ScheduleSpec{"fibonacci", ScheduleAlgorithm::Fibonacci, std::span<const std::uint64_t>{}, "9ca4c0f82671392b5e3ea85be17d614bdf45b1c7fadf6a8162712f813b9db2ec"},
};
inline constexpr std::array<Check, 2> k_kitty_direct_rgb_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"format\":\"rgb\",\"height\":\"1\",\"immutable\":true,\"pixel\":\"ff0000\",\"width\":\"1\"}", "ebe00af4abaccfaf4d4bd3c0fe2bd1513cfa32f5f82893007202d546cc5b7b93"},
};
inline constexpr std::array<Action, 1> k_kitty_direct_rgb_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
};
inline constexpr std::array<InputStep, 1> k_kitty_direct_rgba_chunked_input{
    InputStep{InputOp::LiteralHex, "1b5f47663d33322c733d312c763d312c693d322c6d3d313b2f7741411b5c1b5f47693d322c6d3d303b2f773d3d1b5c", 0, 0ULL},
};
inline constexpr std::array<std::uint64_t, 7> k_kitty_direct_rgba_chunked_terminator_split_cuts{1ULL, 28ULL, 29ULL, 30ULL, 31ULL, 45ULL, 46ULL};
inline constexpr std::array<ScheduleSpec, 1> k_kitty_direct_rgba_chunked_schedules{
    ScheduleSpec{"terminator-split", ScheduleAlgorithm::TerminatorCuts, k_kitty_direct_rgba_chunked_terminator_split_cuts, "18b0c3369388b51424bcc8179dcd9ee1cecad59de79c57a091a7a7c1c39a7b7e"},
};
inline constexpr std::array<Check, 2> k_kitty_direct_rgba_chunked_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"atomicGenerationDelta\":\"1\",\"format\":\"rgba\",\"pixel\":\"ff0000ff\"}", "46d6a726ad90a1983c9bbfb18497cac03370763c93f2119d2716487b05327c5b"},
};
inline constexpr std::array<Action, 1> k_kitty_direct_rgba_chunked_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
};
inline constexpr std::array<InputStep, 1> k_kitty_pixel_boundary_input{
    InputStep{InputOp::LiteralHex, "5253464958545552453a6b697474792d706978656c2d626f756e64617279", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 1> k_kitty_pixel_boundary_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "1635c31bcfe16560870201e84b80bcb77f0f37de7370eaf09a2f862cb9452a24"},
};
inline constexpr std::array<Check, 2> k_kitty_pixel_boundary_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":true,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":true}", "b87b21bd5ea0f04d2da3977bff3aa4f80086ecda8d182daa2b215e12e2f320f4"},
    Check{"state", "{\"boundary\":\"16000000\",\"boundaryAcceptedWhenLedgersFit\":true,\"plusOnePredecodeRejected\":true}", "faa250aad2665507d23e7587ca9ab08bb83ed03a5b915905819f2facae2432d0"},
};
inline constexpr std::array<Action, 2> k_kitty_pixel_boundary_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
    Action{"probe-kitty-pixel-limit", "{\"boundary\":\"16000000\",\"plusOne\":\"16000001\"}", "f58b75c66c23e5aed6fa4ab4b2cf63930d04f7b91ff08fa9182a40d67847bdfd"},
};
inline constexpr std::array<InputStep, 1> k_kitty_png_bomb_input{
    InputStep{InputOp::LiteralHex, "1b5f47663d3130302c693d392c613d543b6956424f5277304b47676f414141414e53556845556741412f2f2f2f4141442f2f2f384941514141414141411b5c", 0, 0ULL},
};
inline constexpr std::array<std::uint64_t, 3> k_kitty_png_bomb_terminator_split_cuts{1ULL, 61ULL, 62ULL};
inline constexpr std::array<ScheduleSpec, 2> k_kitty_png_bomb_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "2264d4a669324972cad5f453b09a934fec98f9229c4f016369db654b8fcffcb2"},
    ScheduleSpec{"terminator-split", ScheduleAlgorithm::TerminatorCuts, k_kitty_png_bomb_terminator_split_cuts, "66f4d3575fae23c5c6eae14fd522008fe66b00fe636d2cbcf14845d1d99025cb"},
};
inline constexpr std::array<Check, 2> k_kitty_png_bomb_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":true,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":true}", "b87b21bd5ea0f04d2da3977bff3aa4f80086ecda8d182daa2b215e12e2f320f4"},
    Check{"state", "{\"decodedAllocationAttempted\":false,\"dimensionsPreflighted\":true,\"rejected\":true}", "8808f6ff21b5b732facf2eb135b7dab6f74243de1875517c09f9d229e05169ee"},
};
inline constexpr std::array<Action, 1> k_kitty_png_bomb_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
};
inline constexpr std::array<InputStep, 1> k_kitty_replacement_transient_input{
    InputStep{InputOp::LiteralHex, "5253464958545552453a6b697474792d7265706c6163656d656e742d7472616e7369656e74", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 1> k_kitty_replacement_transient_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "8dd3d3e00f4500fef7a912a367a29ad00f32f18323d91f2c1c799e40e462b068"},
};
inline constexpr std::array<Check, 2> k_kitty_replacement_transient_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":true,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":true}", "b87b21bd5ea0f04d2da3977bff3aa4f80086ecda8d182daa2b215e12e2f320f4"},
    Check{"state", "{\"atomicFailure\":true,\"oldRetainedUntilPublication\":true,\"simultaneousReservations\":\"old,new,compressed,decoded,conversion,publication\"}", "423fcea29ab4559d23df8d747428e57d57cdff4289b03c2332ff598e5711f971"},
};
inline constexpr std::array<Action, 2> k_kitty_replacement_transient_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
    Action{"probe-kitty-replacement-ledger", "{\"forceFailureAfterNewReservation\":true}", "8398fdff03f2f0c62d4d8ca2acd598c1ad96de507572ba4f039781e37a626b26"},
};
inline constexpr std::array<InputStep, 1> k_kitty_three_layers_crop_input{
    InputStep{InputOp::LiteralHex, "1b5f47663d33322c733d312c763d312c693d332c613d543b2f7741412f773d3d1b5c1b5f47613d702c693d332c703d312c7a3d2d322c783d302c793d302c773d312c683d313b1b5c1b5f47613d702c693d332c703d322c7a3d2d312c783d302c793d302c773d312c683d313b1b5c1b5f47613d702c693d332c703d332c7a3d322c783d302c793d302c773d312c683d313b1b5c", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 1> k_kitty_three_layers_crop_schedules{
    ScheduleSpec{"fibonacci", ScheduleAlgorithm::Fibonacci, std::span<const std::uint64_t>{}, "764bfc46fa891334b04ae911b8831828280d62edd035fc2fd06ef4a27ccf0d09"},
};
inline constexpr std::array<Check, 2> k_kitty_three_layers_crop_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"crop\":\"0,0,1,1\",\"layers\":\"-2,-1,2\",\"placementCount\":\"3\"}", "eaeae12973b241003eed47eada6b02841e2238bd93bf48633ed4a2677fd69538"},
};
inline constexpr std::array<Action, 1> k_kitty_three_layers_crop_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
};
inline constexpr std::array<InputStep, 1> k_kitty_zlib_bomb_input{
    InputStep{InputOp::LiteralHex, "1b5f47663d33322c6f3d7a2c733d3531322c763d3531322c693d31302c613d543b654e727477544542414141417771443154323049583641414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414534444150454141513d3d1b5c", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 2> k_kitty_zlib_bomb_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "eff3077c4b900e410dd0cf3f50474739e097e701641636c3342b56f48ce10967"},
    ScheduleSpec{"bytewise", ScheduleAlgorithm::Bytewise, std::span<const std::uint64_t>{}, "e080fff24175bc55db9f0344eb3c77e9b6bc2bf29d91800366d524819ac25efd"},
};
inline constexpr std::array<Check, 2> k_kitty_zlib_bomb_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":true,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":true}", "b87b21bd5ea0f04d2da3977bff3aa4f80086ecda8d182daa2b215e12e2f320f4"},
    Check{"state", "{\"boundedInflater\":true,\"outputBeyondReservation\":false,\"processAlive\":true}", "dc26c07ff58714594c86509d95bb52e1fd00d7fcaf5d64514afda6aac0b7506c"},
};
inline constexpr std::array<Action, 1> k_kitty_zlib_bomb_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
};
inline constexpr std::array<InputStep, 1> k_malformed_utf8_vt_input{
    InputStep{InputOp::LiteralHex, "f0808080c0afeda0801b5b39393939393939393939393939393939393939393939393939393939393939", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 3> k_malformed_utf8_vt_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "278ac77886ac659e757e0ac9649b04625e9f3e743f300ed7344e96f2a51012b2"},
    ScheduleSpec{"bytewise", ScheduleAlgorithm::Bytewise, std::span<const std::uint64_t>{}, "f05f6ec284f0fd4b589333cce9baa0e94fc7ecc7bf418ca8d8fdd808bfcb52c8"},
    ScheduleSpec{"fibonacci", ScheduleAlgorithm::Fibonacci, std::span<const std::uint64_t>{}, "2939fb498bc529c1cfe64782a178b2004aac59b32dd85f4a7aa2c4b2b210501e"},
};
inline constexpr std::array<Check, 2> k_malformed_utf8_vt_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"deterministic\":true,\"processAlive\":true,\"replacement\":\"fffd\"}", "ec33258cfa5279dea1ece95f5f70e201f5754bd1146fcd9dbab123a3682e953b"},
};
inline constexpr std::array<Action, 1> k_malformed_utf8_vt_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
};
inline constexpr std::array<InputStep, 6> k_nonkitty_generic_boundary_input{
    InputStep{InputOp::LiteralHex, "1b50", 0, 0ULL},
    InputStep{InputOp::RepeatByte, "", 97, 65536ULL},
    InputStep{InputOp::LiteralHex, "1b5c", 0, 0ULL},
    InputStep{InputOp::LiteralHex, "1b50", 0, 0ULL},
    InputStep{InputOp::RepeatByte, "", 98, 65537ULL},
    InputStep{InputOp::LiteralHex, "1b5c5a", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 2> k_nonkitty_generic_boundary_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "f8a2e0b1f4fd8bae213b9fc178a34eb4f7684984297b828f979e41f3b87d150e"},
    ScheduleSpec{"bytewise", ScheduleAlgorithm::Bytewise, std::span<const std::uint64_t>{}, "faba5e7f20a9d8e66a45c37b0e643889573216de8d5e1248e556ea997c8733e5"},
};
inline constexpr std::array<Check, 2> k_nonkitty_generic_boundary_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":true,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":true}", "b87b21bd5ea0f04d2da3977bff3aa4f80086ecda8d182daa2b215e12e2f320f4"},
    Check{"state", "{\"boundary\":\"65536\",\"boundaryAccepted\":true,\"plusOneDiscarded\":true}", "dae72f791640cdaa5c2b0a471ac9a27ece93aa4197e0d328a5b01e3d842c7d9c"},
};
inline constexpr std::array<Action, 2> k_nonkitty_generic_boundary_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
    Action{"probe-generic-escape-limit", "{\"boundary\":\"65536\",\"plusOne\":\"65537\"}", "29de2318e85986ec6d11ad915235a246d6743361f1b0b8dc9179525be2423a5e"},
};
inline constexpr std::array<InputStep, 6> k_osc52_boundaries_input{
    InputStep{InputOp::LiteralHex, "1b5d35323b633b", 0, 0ULL},
    InputStep{InputOp::RepeatByte, "", 65, 1048576ULL},
    InputStep{InputOp::LiteralHex, "07", 0, 0ULL},
    InputStep{InputOp::LiteralHex, "1b5d35323b633b", 0, 0ULL},
    InputStep{InputOp::RepeatByte, "", 65, 1048577ULL},
    InputStep{InputOp::LiteralHex, "07", 0, 0ULL},
};
inline constexpr std::array<std::uint64_t, 5> k_osc52_boundaries_terminator_split_cuts{1ULL, 1048583ULL, 1048584ULL, 1048585ULL, 2097168ULL};
inline constexpr std::array<ScheduleSpec, 2> k_osc52_boundaries_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "a397817e8e3e5e9ee006e5baebc0dc38a57bf75fdf614e36360a35b09457a424"},
    ScheduleSpec{"terminator-split", ScheduleAlgorithm::TerminatorCuts, k_osc52_boundaries_terminator_split_cuts, "7632041b009f014e1afdb99567d4f434cf6f6ff6563869da659910607681a898"},
};
inline constexpr std::array<Check, 2> k_osc52_boundaries_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":true,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":true}", "b87b21bd5ea0f04d2da3977bff3aa4f80086ecda8d182daa2b215e12e2f320f4"},
    Check{"state", "{\"decodedBoundary\":\"786432\",\"encodedBoundary\":\"1048576\",\"partialEffect\":false,\"plusOneDiscarded\":true}", "820d315dba698fdccc92d262d8273699449ec0803ce86bf121572ccc6213ea4a"},
};
inline constexpr std::array<Action, 2> k_osc52_boundaries_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
    Action{"probe-osc52-limits", "{\"decoded\":\"786432\",\"encoded\":\"1048576\"}", "45691a5714c12bdfddf1f05e0d73c335ad36488da5b96e6afff43284e7aa2292"},
};
inline constexpr std::array<InputStep, 1> k_pending_effects_boundary_input{
    InputStep{InputOp::LiteralHex, "5253464958545552453a70656e64696e672d656666656374732d626f756e64617279", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 1> k_pending_effects_boundary_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "488588a392e8437c3e0d6c9078768549d380491f384e32626e0078266145c22e"},
};
inline constexpr std::array<Check, 2> k_pending_effects_boundary_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":true,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":true}", "b87b21bd5ea0f04d2da3977bff3aa4f80086ecda8d182daa2b215e12e2f320f4"},
    Check{"state", "{\"byteBoundary\":\"2097152\",\"descriptorBoundary\":\"256\",\"plusOneAtomic\":true}", "66c83fb86e9722fd14a074e2b2fbf11cd64260cbbd0e72159485cd32d0312f99"},
};
inline constexpr std::array<Action, 2> k_pending_effects_boundary_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
    Action{"probe-pending-effect-ledger", "{\"bytes\":\"2097152\",\"descriptors\":\"256\"}", "d05daaca89cac0a98af063b61ceacb3381e4f002772706604638ad43d307911f"},
};
inline constexpr std::array<InputStep, 1> k_resize_reflow_selection_input{
    InputStep{InputOp::LiteralHex, "6c6f676963616c2d6c696e652d312065cc8120f09f99820d0a6c6f676963616c2d6c696e652d3220776964653ae7958c", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 1> k_resize_reflow_selection_schedules{
    ScheduleSpec{"fibonacci", ScheduleAlgorithm::Fibonacci, std::span<const std::uint64_t>{}, "3295a506eeb7f2db21414df6aef58606b1218df6f54344a965d3f46492b04716"},
};
inline constexpr std::array<Check, 2> k_resize_reflow_selection_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"graphemeStable\":true,\"logicalLines\":\"2\",\"selectionStable\":true,\"wideCellStable\":true}", "8796660c6b9f856960afdbfc4c5e9ae3a93c1ae87c1776416f5e0aeb95f4a7b9"},
};
inline constexpr std::array<Action, 3> k_resize_reflow_selection_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
    Action{"select-logical-range", "{\"end\":\"1:12\",\"start\":\"0:0\"}", "7483bcd8187ea4050596866c46189043a236a3175907df0a16cfa86fed664f8b"},
    Action{"resize-sequence", "{\"dimensions\":\"80x24,20x24,100x24\"}", "35bdf1669109c59a7963e3018a9ed26d1a055a8eb3274de606f55e3000c50cd9"},
};
inline constexpr std::array<InputStep, 1> k_scrollback_byte_boundary_input{
    InputStep{InputOp::RepeatByte, "", 10, 67108865ULL},
};
inline constexpr std::array<ScheduleSpec, 1> k_scrollback_byte_boundary_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "b0af4fa6cc7aa444d6baabcf9079b4e9af4c72ebe5ffc3a5e8226306ceb9fa65"},
};
inline constexpr std::array<Check, 2> k_scrollback_byte_boundary_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":true,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":true}", "b87b21bd5ea0f04d2da3977bff3aa4f80086ecda8d182daa2b215e12e2f320f4"},
    Check{"state", "{\"boundary\":\"67108864\",\"plusOnePreallocationHandled\":true,\"unit\":\"bytes\"}", "32c740ec8455722eecefc697b5c621c494455cfd651cc281a536d0c66bdb1bee"},
};
inline constexpr std::array<Action, 2> k_scrollback_byte_boundary_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
    Action{"probe-terminal-text-ledger", "{\"boundary\":\"67108864\",\"plusOne\":\"67108865\"}", "34db173365c6cd4007113fe2f5a515ba6ae16819d3cebf8b1d6bf48ebcf72ea5"},
};
inline constexpr std::array<InputStep, 1> k_snapshot_dirty_retry_input{
    InputStep{InputOp::LiteralHex, "736e617073686f742d6f6e650d0a736e617073686f742d74776f", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 1> k_snapshot_dirty_retry_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "54a019f6a8ad91ad844969ab7b5aaf0dbfd4cd3597885a7dbd0c3574ae63bdd4"},
};
inline constexpr std::array<Check, 2> k_snapshot_dirty_retry_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"failedPublicationRetried\":true,\"fullSnapshotOnRetry\":true,\"notificationLost\":false}", "8a78317a5f5bece92613c3cc1a0d926035e9cd172b703f5f3ea87fa8e9cca79d"},
};
inline constexpr std::array<Action, 2> k_snapshot_dirty_retry_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
    Action{"snapshot-copy-fail-retry", "{\"failAfterCells\":\"1\",\"requireFullRetry\":true}", "4c194824bb56af651ad184584b8de223e71bc768cf4f816c9f434703a3d91e74"},
};
inline constexpr std::array<InputStep, 1> k_snapshot_mutation_barrier_input{
    InputStep{InputOp::LiteralHex, "626f72726f7765641b5f47663d33322c733d312c763d312c693d31322c613d543b2f7741412f773d3d1b5c", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 1> k_snapshot_mutation_barrier_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "f2911bd22ab4320934f144dc96736b0cfe92463a01cde2674d04a8074590559e"},
};
inline constexpr std::array<Check, 2> k_snapshot_mutation_barrier_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"borrowedAfterEnd\":\"invalid\",\"copiedSnapshotStable\":true,\"mutationBarrier\":true}", "ce8e06018bc3b7799ab1fbeba00ce376afcf37fea0bdb7d2acd4a4d3f4f9291c"},
};
inline constexpr std::array<Action, 2> k_snapshot_mutation_barrier_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
    Action{"snapshot-end-mutate-probe", "{\"mutate\":\"resize-and-write\",\"retainBorrow\":false}", "17b7d6cc397ce5950794874ed65ee5cebcd6e01720264bae83aead75ec9aa778"},
};
inline constexpr std::array<InputStep, 1> k_synchronized_output_input{
    InputStep{InputOp::LiteralHex, "1b5b3f323032366866697273740d0a7365636f6e641b5b3f323032366c", 0, 0ULL},
};
inline constexpr std::array<std::uint64_t, 6> k_synchronized_output_terminator_split_cuts{1ULL, 7ULL, 8ULL, 21ULL, 22ULL, 28ULL};
inline constexpr std::array<ScheduleSpec, 1> k_synchronized_output_schedules{
    ScheduleSpec{"terminator-split", ScheduleAlgorithm::TerminatorCuts, k_synchronized_output_terminator_split_cuts, "6fd3ef2c19725a69903907731cc4707049bde3baec6a8d199ecd3df88b4fcfed"},
};
inline constexpr std::array<Check, 2> k_synchronized_output_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"publishedDuringSync\":false,\"releaseAtomic\":true,\"retainedDirty\":true}", "461c0c1c0c7cb9d362d1ad1c2e6db74fa1f483390a4417f5f9d4573d75e0dc18"},
};
inline constexpr std::array<Action, 1> k_synchronized_output_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
};
inline constexpr std::array<InputStep, 1> k_teardown_callback_quiet_input{
    InputStep{InputOp::LiteralHex, "5253464958545552453a74656172646f776e2d63616c6c6261636b2d7175696574", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 1> k_teardown_callback_quiet_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "af1e5b0206eeaf54c5816612868ce4fe10238d049e8b17ee8a850d8a832773f0"},
};
inline constexpr std::array<Check, 2> k_teardown_callback_quiet_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"callbackAfterDetach\":\"0\",\"producersStoppedFirst\":true,\"quietBeforeUnload\":true}", "30b5223d70416ca4b8322b65152b921a9dc6dedd74f2db1c1c4932a80dfd081a"},
};
inline constexpr std::array<Action, 2> k_teardown_callback_quiet_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
    Action{"teardown-sequence", "{\"order\":\"stop-producers,cancel,stop-posting,detach-callback,release,unload\"}", "afb197e6c56a4b794ac0bf18dcce73847ef204e40c939bd55e282b947606cd9f"},
};
inline constexpr std::array<InputStep, 12> k_title_cwd_boundaries_input{
    InputStep{InputOp::LiteralHex, "1b5d323b", 0, 0ULL},
    InputStep{InputOp::RepeatByte, "", 116, 4096ULL},
    InputStep{InputOp::LiteralHex, "07", 0, 0ULL},
    InputStep{InputOp::LiteralHex, "1b5d323b", 0, 0ULL},
    InputStep{InputOp::RepeatByte, "", 117, 4097ULL},
    InputStep{InputOp::LiteralHex, "07", 0, 0ULL},
    InputStep{InputOp::LiteralHex, "1b5d373b66696c653a2f2f6c6f63616c686f73742f", 0, 0ULL},
    InputStep{InputOp::RepeatByte, "", 99, 131072ULL},
    InputStep{InputOp::LiteralHex, "07", 0, 0ULL},
    InputStep{InputOp::LiteralHex, "1b5d373b66696c653a2f2f6c6f63616c686f73742f", 0, 0ULL},
    InputStep{InputOp::RepeatByte, "", 100, 131073ULL},
    InputStep{InputOp::LiteralHex, "07", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 2> k_title_cwd_boundaries_schedules{
    ScheduleSpec{"all-at-once", ScheduleAlgorithm::AllAtOnce, std::span<const std::uint64_t>{}, "8db64989771edf9ca0e6d9f718357ef5724e275ef91c96b2f5202250533dd305"},
    ScheduleSpec{"bytewise", ScheduleAlgorithm::Bytewise, std::span<const std::uint64_t>{}, "29580905570bf5a9dbf1e4e87e61df4eee514060e04f5153e8ffe5c8046e4c7c"},
};
inline constexpr std::array<Check, 2> k_title_cwd_boundaries_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":true,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":true}", "b87b21bd5ea0f04d2da3977bff3aa4f80086ecda8d182daa2b215e12e2f320f4"},
    Check{"state", "{\"cwdBoundary\":\"131072\",\"plusOneDiscarded\":true,\"titleBoundary\":\"4096\"}", "d95ab76440c115da65ee5378203984be9b8d57a57c2854143126a1ace9def942"},
};
inline constexpr std::array<Action, 2> k_title_cwd_boundaries_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
    Action{"probe-title-cwd-limits", "{\"cwd\":\"131072\",\"title\":\"4096\"}", "e60956a45f465f0cd387577219871d34ff5fb751271747b2296183c9b0cd65b2"},
};
inline constexpr std::array<InputStep, 1> k_unicode_grapheme_wide_input{
    InputStep{InputOp::LiteralHex, "65cc8120e7958c20f09f91a9e2808df09f92bb20e29da4efb88f", 0, 0ULL},
};
inline constexpr std::array<ScheduleSpec, 2> k_unicode_grapheme_wide_schedules{
    ScheduleSpec{"bytewise", ScheduleAlgorithm::Bytewise, std::span<const std::uint64_t>{}, "f3c85a0e685e0b68153b202d0ecf120cb52b71aa91fed40743fa922b15666644"},
    ScheduleSpec{"fibonacci", ScheduleAlgorithm::Fibonacci, std::span<const std::uint64_t>{}, "acb0a0f8e522053a27200088285d4496bd710443c5804d071825f34b7e0a0c4b"},
};
inline constexpr std::array<Check, 2> k_unicode_grapheme_wide_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"combiningStable\":true,\"variationStable\":true,\"wideContinuationStable\":true,\"zwjStable\":true}", "99070e62ba7ddc24fed60084a2a621177bd7c1a5878215673d1085cf68c5e0ef"},
};
inline constexpr std::array<Action, 1> k_unicode_grapheme_wide_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
};
inline constexpr std::array<InputStep, 1> k_write_response_order_input{
    InputStep{InputOp::LiteralHex, "1b5b356e1b5b366e1b5d323b6f7264657265641b5c", 0, 0ULL},
};
inline constexpr std::array<std::uint64_t, 9> k_write_response_order_terminator_split_cuts{1ULL, 3ULL, 4ULL, 5ULL, 7ULL, 8ULL, 9ULL, 19ULL, 20ULL};
inline constexpr std::array<ScheduleSpec, 1> k_write_response_order_schedules{
    ScheduleSpec{"terminator-split", ScheduleAlgorithm::TerminatorCuts, k_write_response_order_terminator_split_cuts, "b9aeb51ca63fcfa35f18ed671821a20d3f3471760fa9c2b97283dc62a60446c7"},
};
inline constexpr std::array<Check, 2> k_write_response_order_checks{
    Check{"resource-trace", "{\"allReservationsReleased\":true,\"checkedArithmetic\":true,\"limitFixture\":false,\"peakWithinFrozenLimits\":true,\"rejectionBeforeAllocation\":false}", "d1dc16dd246375428d05ff5a5de34d1a065ce0f2fc49e6f72996c8d1c497c505"},
    Check{"state", "{\"admittedOrder\":\"protocol-reply,input,resize,side-channel\",\"reordered\":false}", "e005aefd07b0aac008f6fe282090b3391d5b2deacc6a7674e8f251323a2a85b4"},
};
inline constexpr std::array<Action, 2> k_write_response_order_actions{
    Action{"feed-current-schedule", "{\"input\":\"inputPlan\",\"schedule\":\"current\"}", "eee7e3c80e58461cae5046939b65955136b71c321f86fa9c807ef8c86b41e984"},
    Action{"admit-external-events", "{\"order\":\"input,resize,side-channel\"}", "faa34e9f90a6355043b4eb6fb6ee1b984bce394d0e9541808c8c5d4fb7c8404a"},
};

inline constexpr std::string_view kFixtureManifestSha256 = "af7277b42f6e74d147bb5df5781b4e4df03bddfcd28738330dd2403da22b9b44";
inline constexpr std::array<std::string_view, 5> kPerformanceFixtureIds{
    "alternate-screen-roundtrip",
    "color-style-matrix",
    "synchronized-output",
    "unicode-grapheme-wide",
    "write-response-order",
};
inline constexpr PerformancePayloadSpec kPerformancePayload{kPerformanceFixtureIds, 8388608ULL, "ed2fcda3593e3e4053868dec46a26f57981b4f6421361403234437dc9b7a5d1f"};
inline constexpr std::array<Fixture, 29> kFixtures{
    Fixture{"alternate-screen-roundtrip", k_alternate_screen_roundtrip_input, 32ULL, "6690d5ef558d785073e4d2b43291332cb108a06bf09245e671bd3521421c5e4c", k_alternate_screen_roundtrip_schedules, k_alternate_screen_roundtrip_actions, k_alternate_screen_roundtrip_checks, "68eed2ff76fe88b47de895ddeea7ceee9e1ce7050aa05219f08459478b0ed491"},
    Fixture{"bracketed-paste-modes", k_bracketed_paste_modes_input, 8ULL, "dfe5456fed7d6508bcf91708c48eb9d48213ca34668e6238e94cace519f39755", k_bracketed_paste_modes_schedules, k_bracketed_paste_modes_actions, k_bracketed_paste_modes_checks, "83be78b30ccde82b52914b4e4434ff844da89c374dcf2a856c4f2ada4d998e61"},
    Fixture{"color-style-matrix", k_color_style_matrix_input, 61ULL, "bbc2be22c7efb8045189d6a90472be9fc4c08c68361c4fb131974502369074a9", k_color_style_matrix_schedules, k_color_style_matrix_actions, k_color_style_matrix_checks, "692d1b10949b2bad4fcf58e27999f041c9a09d6946e469d9fa77af383f955788"},
    Fixture{"cursor-and-palette", k_cursor_and_palette_input, 51ULL, "159aeb95d9b82ee8c5aca7c41ff3c3f84c5178b9a7717f11f1bb0e77944c8a9d", k_cursor_and_palette_schedules, k_cursor_and_palette_actions, k_cursor_and_palette_checks, "fc6ec4ad975d0682244efdfa8fcd6a65d8a013d3ddd0284020714bd38fce51e8"},
    Fixture{"focus-and-mouse-matrix", k_focus_and_mouse_matrix_input, 64ULL, "d2bfbec5345d3d5d1f3a7fa63eabb2606412c7a33e991d9b921475511f072dfa", k_focus_and_mouse_matrix_schedules, k_focus_and_mouse_matrix_actions, k_focus_and_mouse_matrix_checks, "c70374b9e82a9a4cc29c1672dc2d3a241e5153b7fdba219987c5a16437b53093"},
    Fixture{"hyperlink-boundaries", k_hyperlink_boundaries_input, 18468ULL, "2eb460fdd6803c1287a0b6dbcc66ac9e4c4383689561d022f43025dd51bb3430", k_hyperlink_boundaries_schedules, k_hyperlink_boundaries_actions, k_hyperlink_boundaries_checks, "6fa78b96873110c3dcf3b0dd486f6d0f39c65bea6114b3c06751c38debfa9dcc"},
    Fixture{"kitty-delete-generation", k_kitty_delete_generation_input, 73ULL, "dff5c6687034aab2d6ccc01f584f12fedffae9d46ca6fa338e3ee2c61b408fec", k_kitty_delete_generation_schedules, k_kitty_delete_generation_actions, k_kitty_delete_generation_checks, "5c615cdff7650f746f7f5f7c1ec3e23ff6fa84181423b55ca6d45ca2df5b6952"},
    Fixture{"kitty-dimension-overflow", k_kitty_dimension_overflow_input, 48ULL, "6bf0e1c72245de0eb8009394a577fceb2fe1cb2b6924fda955be6cc885502c4c", k_kitty_dimension_overflow_schedules, k_kitty_dimension_overflow_actions, k_kitty_dimension_overflow_checks, "6bec9f1f72f59159992dbdb0de1851d349b8273f5c32f6fb43aaf0591a2d4a2c"},
    Fixture{"kitty-direct-png", k_kitty_direct_png_input, 116ULL, "a90a604995b9f736500e5354d9ced001674c4a26aad8b1353ceba3792986cf39", k_kitty_direct_png_schedules, k_kitty_direct_png_actions, k_kitty_direct_png_checks, "ce9bc0f53831b6be3c62597157607c016e82aaec5d23689f24076d8ac879a00f"},
    Fixture{"kitty-direct-rgb", k_kitty_direct_rgb_input, 30ULL, "c773307ded70943d7ad8e46b1ebb5277cf7a273d4cdcc231e0e71884f8a1d8c5", k_kitty_direct_rgb_schedules, k_kitty_direct_rgb_actions, k_kitty_direct_rgb_checks, "94ec0a32309d79d7b3a430ea4be006c2f0af64e1143a2e791c699c04c1b588f4"},
    Fixture{"kitty-direct-rgba-chunked", k_kitty_direct_rgba_chunked_input, 47ULL, "1af3a59a5690c06443118d263de9ec61759f1e8f005d47bf06d5fa9ced9400b3", k_kitty_direct_rgba_chunked_schedules, k_kitty_direct_rgba_chunked_actions, k_kitty_direct_rgba_chunked_checks, "8360b87bba715185f3049cbc4b1a1176edf117a39b87905dc010a21c0fbe6081"},
    Fixture{"kitty-pixel-boundary", k_kitty_pixel_boundary_input, 30ULL, "d3c18380244421460d49a91993a94b7adf34c3c3180fab3fdc2ad8a3aa21663c", k_kitty_pixel_boundary_schedules, k_kitty_pixel_boundary_actions, k_kitty_pixel_boundary_checks, "0171f02fa149d8f03e5236b22fc334d5ee43f8eee6f8ab08388faaf4882c872a"},
    Fixture{"kitty-png-bomb", k_kitty_png_bomb_input, 63ULL, "38788941c8f81169d481e25a646da2ffd8657d7fecc1feb04703e32c1469c143", k_kitty_png_bomb_schedules, k_kitty_png_bomb_actions, k_kitty_png_bomb_checks, "cef99b4de124a9bf08c06e117643f0a4e17a9c6401edbbcb596ae139f088f8f7"},
    Fixture{"kitty-replacement-transient", k_kitty_replacement_transient_input, 37ULL, "a0bd005a45a0bf66d3522e2870377c5a764ab867b20f9e5fed7b41634f44f65b", k_kitty_replacement_transient_schedules, k_kitty_replacement_transient_actions, k_kitty_replacement_transient_checks, "cc395d2e068f33e5217c9fdec6a143882ce5b2f6bc99f2d55cecb034e8cbb693"},
    Fixture{"kitty-three-layers-crop", k_kitty_three_layers_crop_input, 147ULL, "216735365bdd9191bcdf2c162c5c692f08f621af14fafcbe098c8b460cea0de3", k_kitty_three_layers_crop_schedules, k_kitty_three_layers_crop_actions, k_kitty_three_layers_crop_checks, "9088cd5c73d2d0c25ab4297de8dc5420fa3eda034196ed59f70eb7d35c1db6a5"},
    Fixture{"kitty-zlib-bomb", k_kitty_zlib_bomb_input, 1423ULL, "3e00f64d79a2108340a3f49571bcc6ad9fec714d76a3029b4541001e4bcc2637", k_kitty_zlib_bomb_schedules, k_kitty_zlib_bomb_actions, k_kitty_zlib_bomb_checks, "4ca9e822e12118a199a3df1b1849eec43480320119334d5fc9a2d901d24c2aac"},
    Fixture{"malformed-utf8-vt", k_malformed_utf8_vt_input, 42ULL, "e4d4872f996a5d1907ce2e1c3b44a9a5f21e28a739b68ef9f32a06751cf70883", k_malformed_utf8_vt_schedules, k_malformed_utf8_vt_actions, k_malformed_utf8_vt_checks, "600dc095f3ce4f608c47d177d340598db4dd65e720255855f8457326b30b3a3e"},
    Fixture{"nonkitty-generic-boundary", k_nonkitty_generic_boundary_input, 131082ULL, "1aa6a75e6774b1ca4bbce8e15b8c4a7d04a400b87c6816baa565f4952124dd84", k_nonkitty_generic_boundary_schedules, k_nonkitty_generic_boundary_actions, k_nonkitty_generic_boundary_checks, "d03b13f029afa260500b87ecfbf4ce92437604c382df1704e5a61ff4c9d91552"},
    Fixture{"osc52-boundaries", k_osc52_boundaries_input, 2097169ULL, "a2cdeef8b6ac68e4353bf945015dbc73755d702539b40e72dd6a8191a3abcd60", k_osc52_boundaries_schedules, k_osc52_boundaries_actions, k_osc52_boundaries_checks, "3146a83c2ebe1430bd3e7e6bdd7abec10edf6c3bd461c7ba8fd7644982381b3a"},
    Fixture{"pending-effects-boundary", k_pending_effects_boundary_input, 34ULL, "7d37ea0801548f97ef90d79013ff56b0109c85857c92408242da43faf7e11421", k_pending_effects_boundary_schedules, k_pending_effects_boundary_actions, k_pending_effects_boundary_checks, "36b85b143a637ba320b77885f375a68b9b70428ae3e2dae12000ba033363bd74"},
    Fixture{"resize-reflow-selection", k_resize_reflow_selection_input, 48ULL, "ca742d13cb3783241719471d2bf390d61d088271a4948176a86edfd83c1ca284", k_resize_reflow_selection_schedules, k_resize_reflow_selection_actions, k_resize_reflow_selection_checks, "aaeee69ecf619601c73c887c3c3b1507cbb79a05fcbc6f1aad8c61efd11abac0"},
    Fixture{"scrollback-byte-boundary", k_scrollback_byte_boundary_input, 67108865ULL, "b83e67934e14d84d649d1a3a5ba68b24999fdc24ced07226f244dcb1d8bbd888", k_scrollback_byte_boundary_schedules, k_scrollback_byte_boundary_actions, k_scrollback_byte_boundary_checks, "e98d68a04399454b373c45a65a9d0c4f11b58c29c2510f10703cac8f5d567dcb"},
    Fixture{"snapshot-dirty-retry", k_snapshot_dirty_retry_input, 26ULL, "85b130c73c00e2c638f674646a3f1558d8575a46ea74a1bdaea56376512c0e92", k_snapshot_dirty_retry_schedules, k_snapshot_dirty_retry_actions, k_snapshot_dirty_retry_checks, "5d229dda331038ffaf1bf0a410706074e0810dc45035a8985e5010e50cb031e0"},
    Fixture{"snapshot-mutation-barrier", k_snapshot_mutation_barrier_input, 43ULL, "f9f35d32a6682ac83667dee40754d5b10deeb0a5bd404af0fda90be41bf02bdc", k_snapshot_mutation_barrier_schedules, k_snapshot_mutation_barrier_actions, k_snapshot_mutation_barrier_checks, "094a267b142b355f2f3f262939dd135daf0ce8c2c567346b2f850c79c297941d"},
    Fixture{"synchronized-output", k_synchronized_output_input, 29ULL, "63df7fae48ddf50fd90c3afa0acad0fa80e98f55b55e53f9ac35af089f41085a", k_synchronized_output_schedules, k_synchronized_output_actions, k_synchronized_output_checks, "66ce999768a96edfddb51a5479b61ef7bb11be3cd225aaccc12f9d0b4f6f82f8"},
    Fixture{"teardown-callback-quiet", k_teardown_callback_quiet_input, 33ULL, "9f7d2a233f5cb47f96bc48313d4f0cad1bbf109a817f008d4be35b3b9d4ffef1", k_teardown_callback_quiet_schedules, k_teardown_callback_quiet_actions, k_teardown_callback_quiet_checks, "f287bdb487c682995ae377d449100ffc1db2e1f245de3c72b305b0bbf6721c2b"},
    Fixture{"title-cwd-boundaries", k_title_cwd_boundaries_input, 270392ULL, "9e9df92a42db277fd6cf01e220186f05bc4a7ba8d4abf07187da46fcee237aa3", k_title_cwd_boundaries_schedules, k_title_cwd_boundaries_actions, k_title_cwd_boundaries_checks, "97bb5ac725de2a2a3ad605a18d0dff7adcd3b675c5eac50f8ab5f14363477769"},
    Fixture{"unicode-grapheme-wide", k_unicode_grapheme_wide_input, 26ULL, "9734cebf077f2bdb7d9752540ff01999a74b7a5e13235a5e275ee4b19ad64919", k_unicode_grapheme_wide_schedules, k_unicode_grapheme_wide_actions, k_unicode_grapheme_wide_checks, "5d139cf5d387a81ab61747e78641c9a8f2f9b02694f4a9316e493a73f5436a9a"},
    Fixture{"write-response-order", k_write_response_order_input, 21ULL, "276e64c482e03917f0d2d3a4b76ce198f55979917cf0d92afa95a4996b03143f", k_write_response_order_schedules, k_write_response_order_actions, k_write_response_order_checks, "ce9e649789ca457180bde023014aaba2d03d595b7350247319ee9c387c5cbf65"},
};
[[nodiscard]] inline constexpr std::span<const Fixture> GetFixtures() noexcept { return kFixtures; }
} // namespace RedSalamander::TerminalEngine::Fixtures::V1

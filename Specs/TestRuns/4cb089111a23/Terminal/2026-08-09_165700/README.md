# Terminal Release localization and routed WindowProc proof

- Status: passed
- Build: Release x64 with `ENABLE_TESTS`
- Scenario: `terminal_release_localization_routing_refactor`
- Command: `.build\x64\Release\PluginContractTests.exe --terminal-selftests`
- Result: 21 live localization/real-object assertions and 45 deterministic engine assertions passed; the Terminal module reached its unload quiet point.
- Coverage: invariant JSON schema fields across English, French, and Japanese; translated schema display text; live UI Automation Name; localized runtime-loader failure through a real `ITerminal`; the routed create/teardown lifecycle; maximum-size Kitty conversion; replacement, cancellation, stale-generation rejection, and unload behavior.
- Performance result: the deterministic 64 MiB fixture retained its 128 MiB resident cap and one-worker/replacement policy. Maximum conversion was 28,834 microseconds versus 29,138 microseconds in the Package 10 baseline on the same machine and fixture.
- Build evidence: x64 `.build/logs/msbuild-20260809_165504_419.log` and `.build/logs/msbuild-20260809_165612_035.log`, plus ARM64 `.build/logs/msbuild-20260809_165745_844.log`; all completed with zero warnings and zero errors.

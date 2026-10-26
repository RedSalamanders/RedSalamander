# RedSalamander C++ compilation and linkage audit

Date: 2026-10-23. All tracked vcxproj files, x64/ARM64, Debug/Release/ASan Debug; effective MSBuild evaluation with full compressed metadata retained. This is metadata evidence; native build and runtime qualification are separate.

Both baseline and candidate select installed default MSVC 14.52.37025 with v145 and SDK 10.0.26100.0. An older toolset is unavailable locally. The intended experiment changes Release optimization policy, not compiler version or the adopted DxUi API pin.

Release explicitly requires /Gy /GL and final /OPT:REF /LTCG, with ordinary incremental linking disabled. Debug and ASan diagnostic optimization remain unchanged. Root late imports are intentional: setting WPO too early permits the MSVC delayed WPO import to restore incremental LTCG.

## C++ settings after implementation

The table includes every distinct effective value of each compiler field across all projects and both platforms. all-settings.csv.gz maps every value to its exact project/configuration/platform/files and includes all linker, librarian and property fields on both sides. Compressed settings-before/after.json.gz preserve full paths, project references and source inventories.

| Setting | Debug | Release | ASan Debug |
| --- | --- | --- | --- |
| AdditionalOptions | /FS ; /utf-8 /external:anglebrackets /external:templates- /wd5045 /wd4820 /bigobj /MP /FS ; /utf-8 /external:anglebrackets /external:templates- /wd5045 /wd4820 /FS ; /utf-8 /external:anglebrackets /external:templates- /wd5045 /wd4820 /MP /FS ; /utf-8 /external:anglebrackets /external:templates- /wd5045 /wd4820 /utf-8 /external:anglebrackets /external:templates- /wd5045 /wd4820 /FS ; /utf-8 /FS  | /FS ; /utf-8 /external:anglebrackets /external:templates- /wd5045 /wd4820 /bigobj /MP /FS ; /utf-8 /external:anglebrackets /external:templates- /wd5045 /wd4820 /FS ; /utf-8 /external:anglebrackets /external:templates- /wd5045 /wd4820 /MP /FS ; /utf-8 /FS  | /FS ; /fsanitize=address /utf-8 /FS ; /utf-8 /external:anglebrackets /external:templates- /wd5045 /wd4820 /bigobj /MP /FS ; /utf-8 /external:anglebrackets /external:templates- /wd5045 /wd4820 /FS ; /utf-8 /external:anglebrackets /external:templates- /wd5045 /wd4820 /MP /FS ; /utf-8 /external:anglebrackets /external:templates- /wd5045 /wd4820 /utf-8 /external:anglebrackets /external:templates- /wd5045 /wd4820 /FS ; /utf-8 /FS  |
| BasicRuntimeChecks | EnableFastChecks | Default | Default |
| BufferSecurityCheck | true | true | true |
| ConformanceMode | true | true | true |
| DebugInformationFormat | EditAndContinue | ProgramDatabase | ProgramDatabase |
| EnableASAN | false | false | true |
| EnableEnhancedInstructionSet | NotSet | NotSet | NotSet |
| ExceptionHandling | Sync | Sync | Sync |
| ExternalWarningLevel | InheritWarningLevel; TurnOffAllWarnings | InheritWarningLevel; TurnOffAllWarnings | InheritWarningLevel; TurnOffAllWarnings |
| FavorSizeOrSpeed | Neither | Neither | Neither |
| FloatingPointModel | Precise | Precise | Precise |
| FunctionLevelLinking | <unset> | true | <unset> |
| IntrinsicFunctions | false | false; true | false |
| LanguageStandard | stdcpplatest | stdcpplatest | stdcpplatest |
| MultiProcessorCompilation | <unset>; true | <unset>; true | <unset>; true |
| OmitFramePointers | <unset>; false | <unset>; false | <unset>; false |
| Optimization | Disabled | MaxSpeed | Disabled |
| PrecompiledHeader | ; Create; NotUsing; Use | ; Create; NotUsing; Use | ; Create; NotUsing; Use |
| RuntimeLibrary | MultiThreadedDebug; MultiThreadedDebugDll; MultiThreadedDebugDLL | MultiThreaded; MultiThreadedDll; MultiThreadedDLL | MultiThreadedDebugDLL |
| SDLCheck | true | true | true |
| TreatWarningAsError | false; true | false; true | false; true |
| WarningLevel | EnableAllWarnings; Level3; Level4 | EnableAllWarnings; Level3; Level4 | EnableAllWarnings; Level3; Level4 |
| WholeProgramOptimization | false | true | false |

Definitions, warning-number suppressions and include search paths are file-specific and retained verbatim in the attached data. They are not flattened into an assertion of identical build behavior.

## Link and archive settings after implementation

| Setting | Debug | Release | ASan Debug |
| --- | --- | --- | --- |
| OptimizeReferences | <unset> | true | <unset> |
| LinkTimeCodeGeneration | <unset> | UseLinkTimeCodeGeneration | <unset> |
| EnableCOMDATFolding | <unset> | true | <unset> |
| GenerateDebugInformation | DebugFull; false; true | DebugFull; false; true | DebugFull; false; true |
| SubSystem | Console; Windows | Console; Windows | Console; Windows |
| TargetMachine | MachineARM64; MachineX64 | MachineARM64; MachineX64 | MachineARM64; MachineX64 |
| DataExecutionPrevention | true | true | true |
| RandomizedBaseAddress | true | true | true |
| HighEntropyVA | <unset> | <unset> | <unset> |
| ControlFlowGuard | <unset> | <unset> | <unset> |
| AdditionalOptions | <unset> | <unset> | <unset> |
| IgnoreAllDefaultLibraries | <unset> | <unset> | <unset> |
| AdditionalDependencies | See exact per-project CSV | See exact per-project CSV | See exact per-project CSV |

A .lib is an archive; /OPT:REF removes unreferenced COMDATs at final EXE/DLL links. LTO cannot optimize through an independent DLL boundary. Referenced factories/exports and /INCLUDE roots can retain otherwise unused functionality. No claim is made that every unused DxUi control disappears.

## Interpretation and sources

Unset Gy metadata can still mean /Gy through /O1 or /O2. Unset REF with /DEBUG instead defaults to NOREF. Existing ICF, warning, ISA, exception, size/speed, CRT and security choices remain project-specific. Runtime speed and final size require measured results; larger GL archives are expected to carry optimizer data.

Microsoft references: [GL](https://learn.microsoft.com/en-us/cpp/build/reference/gl-whole-program-optimization), [LTCG](https://learn.microsoft.com/en-us/cpp/build/reference/ltcg-link-time-code-generation), [OPT](https://learn.microsoft.com/en-us/cpp/build/reference/opt-optimizations), [exact tool compatibility](https://learn.microsoft.com/en-us/cpp/porting/binary-compat-2015-2017).

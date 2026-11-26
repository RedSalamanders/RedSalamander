# RedSalamander performance and resource contract

Status: current normative product contract
Owner: RedSalamander product and test harness maintainers

## Mandate

Design, implementation and review must preserve correctness and minimize CPU time, allocations, retained memory, copies, synchronization, GPU submissions and wake-ups. Prefer the lower steady-state cost among correct designs. Keep queues, caches and retained work within the limits set by the owning domain. Apply the owning domain's wake-up, visibility and lifetime rules when work settles or a view becomes inactive.

Changes that can affect responsiveness, startup, traversal, rendering, search, file operations, plug-in I/O, background work or retained memory need a named scenario, instrumentation and deterministic coverage from the start. The domain contracts identify the work and metrics to measure.

## Budget authorities

| Scope | Owning contract and measurements |
| --- | --- |
| Startup and initial input | [Startup bootstrap](Core_StartupBootstrap.md) and the named startup scenarios in [performance validation](../Testing/Testing_PerformanceValidation.md). |
| FolderView enumeration, layout and rendering | [FolderView performance contract](../UI/UI_FolderView.md#performance-validation-contract) and the machine-keyed [FolderView budgets](../Testing/FolderViewPerfBudgets.json5). |
| Search and Compare Directories | [Search performance contract](Core_Search.md#performance-validation-contract) and [Compare Directories](Core_CompareDirectories.md). |
| File Operations and provider I/O | [File Operations](../FileSystem/FileSystem_FileOperations.md) and [performance validation](../Testing/Testing_PerformanceValidation.md#overview), including topology, queueing, verification traffic and cancellation. |

## Measurement and acceptance

For performance comparisons, use the same machine, configuration and deterministic scenario for the baseline and candidate. Record the source identity, execution profile and inputs. Keep measured work separate from setup and waiting as described by the owning scenario. Establish or change a hard budget only with stable paired observations and the documented quality checks.

[Performance validation](../Testing/Testing_PerformanceValidation.md) owns test procedures, metric definitions and report interpretation. Archive source-bound evidence under the [validation evidence contract](../Testing/Testing_RunEvidence.md). Record intentional trade-offs with measured results and the affected scope in the owning contract or active plan.

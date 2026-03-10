# SelfTest Specification

## Overview

RedSalamander ships three debug self-test suites:
- `--compare-selftest`
- `--commands-selftest`
- `--fileops-selftest`

This document defines the normative result contract shared by those suites.

Related documents:
- `Specs/TestRuns/README.md`
- `Specs/Testing/Testing_SelfTestRemoteCredentials.md`

## Result Contract

### Case coverage

Every declared self-test case must produce exactly one case result in the suite `results.json`.

Allowed statuses are:
- `passed`
- `failed`
- `skipped`

There must be no declared case with a missing status.

### Skip semantics

- `skipped` is the correct result when a declared case cannot run because a required precondition is absent.
- Every skipped case must carry a human-readable reason.
- Skipping is part of normal suite behavior for conditional coverage and must not make the suite fail by itself.

### Setup failure behavior

If a suite encounters a fatal setup failure before all declared cases can run:
- the setup failure itself must be recorded explicitly,
- unreached declared cases must still be emitted as `skipped`,
- each backfilled skipped case must explain that it was not executed because of suite setup failure.

## Conditional Coverage Rules

Conditional coverage must remain part of the suite membership. Preconditions change execution status, not case existence.

Examples:
- remote smoke tests skip when required connection profiles, secrets, or sandbox roots are absent,
- plugin-dependent tests skip when the plugin is not available,
- machine-dependent filesystem coverage such as ReFS skips when the required volume is not present.

A case must not disappear from the suite just because the current machine lacks its prerequisites.

Environment variables may select alternate test inputs such as profile names, but they must not be used to remove a declared case from suite membership.

## Suite and Aggregated Results

- Suite `results.json` files must preserve per-case status and reason.
- Aggregated self-test results must count `passed`, `failed`, and `skipped` consistently with the suite artifacts.
- Checked-in archived runs may contain skipped cases; that is valid when the skip reason documents the missing precondition.

## Artifact Contract

Self-test artifacts must preserve enough detail to explain why a run passed, failed, or skipped:
- `results.json` records final case status and reason,
- `trace.txt` records supporting diagnostic context,
- archived copies under `Specs/TestRuns/` must keep those files intact.

## Search-Specific Coverage

Search coverage follows the same contract:
- local, fallback, indexed, and service search cases stay declared,
- ReFS validation stays declared even on machines without ReFS,
- a machine without a fixed ReFS volume records `skipped` with a reason instead of silently omitting the case.

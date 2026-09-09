# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

**All development guidelines live in [AGENTS.md](AGENTS.md)** — project overview, architecture, critical rules (RAII, perf validation, spec closeout, tooling governance), regression guards, and the skills table pointing into `.github/skills/`. Read it first; it is the canonical agent guidance for this repository.

## Quick Reference

```powershell
# Build (Debug by default); output: .build\<Platform>\<Configuration>\
.\build.ps1
.\build.ps1 -Configuration Release
.\build.ps1 -ProjectName RedSalamander

# Full local test suite (see README "Self-tests")
.\Tools\Run-AllTests.ps1 -Suite Full
```

# Observatory Track 19 RedConfigure evidence

Debug x64 `RedConfigureTests.exe` passed after the typed preview/apply, invariant numeric grammar,
duplicate-theme, validation, and resource-localization changes.

The same-machine five-sample candidate median was **301 ms** for the 6-owner/1,500-row/10-theme scan,
**3 ms** for validation, **301 us** for a single live theme edit, and **57,111 us** for a validated
512-token mass preview. All remain below the deterministic **500 ms**, **250 ms**, **16 ms**, and
**100 ms** budgets respectively.

The immediately preceding same-machine Debug baseline had one sample (**309 ms / 4 ms / 302 us**),
so the single-edit comparison is directional rather than statistical; it shows no material regression.
The 512-token mass-preview timing was introduced with this change and is compared to its explicit-action
budget rather than to a historical sample.

Build proof: `.build/logs/msbuild-20260717_155717_097.log`, 0 warnings / 0 errors.
